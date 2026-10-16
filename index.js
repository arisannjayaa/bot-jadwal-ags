const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const { google } = require("googleapis");
const fs = require("fs").promises;
const readline = require("readline");
const XLSX = require("xlsx");
const cron = require("node-cron");
const https = require("https");

// ============================================================================
// 1. KONFIGURASI UTAMA
// ============================================================================
const SPREADSHEET_ID = "1DLcMkga8UiRtRJ3ZQIPMRQb-5d1IFiu_"; // Excel Jadwal
const SPREADSHEET_ID_GAJI = "19QJEZ11K63KYfGfiT4ECSHodfsa_KgCpPz7gDg3UWxk"; // Excel Gaji pribadi (JAYAK)
const SPREADSHEET_ID_GAJI_ALL = "1jaRRnWcriFDN-Aar-6Z9sAVVkqnxYzX786QToY26x1I"; // Excel Payroll SELURUH karyawan
const NAMA_SAYA_DI_ABSEN = "JAYAK"; // Sesuai tulisan di absen admin

const ID_TUJUAN_NOTIFIKASI = "628970282769@c.us";
const WAKTU_RONDA_MS = 1 * 60 * 1000; // 1 Menit

// ============================================================================
// KONFIGURASI THROTTLE, RATE LIMIT & RELIABILITAS
// ============================================================================
const JEDA_MINIMAL_NOTIF_MS = 3 * 60 * 1000; // Jeda minimal antar notifikasi revisi (3 menit)
const JEDA_ANTAR_PESAN_MS = 1200; // Jeda antar pesan WA terkirim, biar gak kena flag spam
const MAX_RETRY_RECONNECT = 3; // Percobaan reconnect otomatis sebelum proses di-restart PM2

// Isi 2 baris ini kalau mau notifikasi fallback ke Telegram saat WA down.
// Kosongkan (biarkan string kosong) kalau belum mau pakai fitur ini.
const TELEGRAM_BOT_TOKEN = ""; // Token dari @BotFather
const TELEGRAM_CHAT_ID = ""; // Chat ID tujuan notifikasi fallback

// ============================================================================
// KONFIGURASI RONDA & PENGINGAT PAYROLL
// ============================================================================
const WAKTU_RONDA_GAJI_MS = 5 * 60 * 1000;      // Cek spreadsheet payroll tiap 5 menit
const JEDA_TENANG_GAJI_MS = 10 * 60 * 1000;     // Tunggu data "tenang" 10 menit (admin selesai ngetik) baru kirim notif
const BATAS_TUNGGU_GAJI_MS = 60 * 60 * 1000;    // Maksimal nunggu 1 jam, setelah itu tetap dikirim
const NOTIF_PAYROLL_SEMUA_KARYAWAN = true;      // false = notif cuma soal gaji pribadi (matikan ringkasan karyawan lain)
const HARI_PENGINGAT_PERIODE_1 = 15;            // Tanggal pengingat pencairan periode 1 (periode 2 = hari terakhir bulan)
const FILE_SNAPSHOT_GAJI = "snapshot_gaji.json"; // Disimpan di disk supaya restart bot tidak kehilangan baseline
// Tulisan di sheet (baris TOTAL TRANSFER, kolom kanan) yang berarti periode sudah ditransfer
const PENANDA_GAJI_CAIR = /^(done|lunas|paid|selesai|sudah transfer|sudah tf|transferred|cair)$/i;

// ============================================================================
// KONFIGURASI SIKLUS KONTRAK
// ============================================================================
const BULAN_MULAI_KONTRAK = 10; // 10 = Oktober. Ganti kalau kontrak beda tanggal mulai.

const BULAN_MAP = {
    "JANUARY": 1, "JANUARI": 1, "FEBRUARI": 2, "MARET": 3, "APRIL": 4,
    "MEI": 5, "JUNI": 6, "JULI": 7, "AGUSTUS": 8, "SEPTEMBER": 9,
    "OKTOBER": 10, "NOVEMBER": 11, "DESEMBER": 12
};
const NAMA_BULAN_ID = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

let objekDataLama = null; 
let globalAuthClient = null;
let isBotReady = false;

// State untuk throttle notifikasi revisi
let antrianPerubahanPending = [];
let waktuNotifTerakhir = 0;

// State untuk rate limit pengiriman pesan WA (antrian serial)
let antrianKirimWA = Promise.resolve();

// State untuk retry reconnect otomatis
let jumlahRetryReconnect = 0;

// Cache hasil cek konflik jadwal terakhir, buat dedupe notifikasi
let daftarKonflikSebelumnya = [];

// State ronda payroll
let snapshotGajiDinotif = null;     // kondisi terakhir yang SUDAH diinfokan ke WA
let snapshotGajiTerakhir = null;    // snapshot ronda sebelumnya (buat deteksi kapan data terakhir berubah)
let waktuGajiBerubahTerakhir = 0;
let waktuGajiPendingMulai = 0;
let rondaGajiBerjalan = false;      // cegah 2 ronda payroll jalan bersamaan
let gagalRondaGajiBeruntun = 0;
let rondaJadwalBerjalan = false;    // cegah 2 ronda jadwal jalan bersamaan
let penjadwalSudahAktif = false;    // cegah setInterval/cron dobel kalau event "ready" kepanggil lagi setelah reconnect

// ============================================================================
// 2. SISTEM LOGIN OAUTH 2.0
// ============================================================================
async function authorize() {
    if (globalAuthClient) return globalAuthClient;
    let content;
    try {
        content = await fs.readFile("credentials.json");
    } catch (err) {
        sendSystemAlert("❌ Gagal Login Google Drive: File credentials.json tidak ditemukan!");
        return null;
    }
    const credentials = JSON.parse(content);
    const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    try {
        const token = await fs.readFile("token.json");
        oAuth2Client.setCredentials(JSON.parse(token));
        globalAuthClient = oAuth2Client;
        return oAuth2Client;
    } catch (err) {
        return await getNewToken(oAuth2Client);
    }
}

async function getNewToken(oAuth2Client) {
    const authUrl = oAuth2Client.generateAuthUrl({ access_type: "offline", scope: ["https://www.googleapis.com/auth/drive.readonly"] });
    console.log("\n=========================================\nBuka link ini:\n" + authUrl + "\n=========================================\n");
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    return new Promise((resolve) => {
        rl.question("Paste kode di sini: ", async (code) => {
            rl.close();
            const { tokens } = await oAuth2Client.getToken(code);
            oAuth2Client.setCredentials(tokens);
            await fs.writeFile("token.json", JSON.stringify(tokens));
            globalAuthClient = oAuth2Client;
            resolve(oAuth2Client);
        });
    });
}

// Parse nama sheet ("JULI 2026", "JANUARY 2026") jadi { bulanNum, tahun }
function parseSheetBulanTahun(sheetName) {
    const upper = sheetName.toUpperCase();
    let bulanNum = null;
    for (const key in BULAN_MAP) {
        if (upper.includes(key)) { bulanNum = BULAN_MAP[key]; break; }
    }
    const tahunMatch = sheetName.match(/\d{4}/);
    const tahunSheetIni = tahunMatch ? parseInt(tahunMatch[0]) : new Date().getFullYear();
    return { bulanNum, tahun: tahunSheetIni };
}

// Cek apakah sebuah sheet termasuk dalam window kontrak tertentu
function sheetMasukKontrak(sheetBulan, sheetTahun, periode) {
    if (sheetTahun === periode.tahunMulai && sheetBulan >= BULAN_MULAI_KONTRAK) return true;
    if (sheetTahun === periode.tahunSelesai && sheetBulan < BULAN_MULAI_KONTRAK) return true;
    return false;
}

// Hitung window siklus kontrak yang sedang berjalan, relatif ke bulan/tahun sheet aktif
function getPeriodeKontrak(bulanAktifNum, tahunAktif) {
    const tahunMulai = bulanAktifNum >= BULAN_MULAI_KONTRAK ? tahunAktif : tahunAktif - 1;
    const tahunSelesai = tahunMulai + 1;
    const bulanKe = bulanAktifNum >= BULAN_MULAI_KONTRAK
        ? (bulanAktifNum - BULAN_MULAI_KONTRAK + 1)
        : (bulanAktifNum + (12 - BULAN_MULAI_KONTRAK) + 1);
    return {
        tahunMulai, tahunSelesai, bulanKe,
        label: `${NAMA_BULAN_ID[BULAN_MULAI_KONTRAK]} ${tahunMulai} – ${NAMA_BULAN_ID[BULAN_MULAI_KONTRAK - 1] || "September"} ${tahunSelesai}`
    };
}

// ============================================================================
// 3. FUNGSI BANTUAN & FORMATTING
// ============================================================================
function formatTanggalExcel(val) {
    if (!val) return "-";
    if (!isNaN(val) && val > 40000) {
        const date = new Date(Math.round((val - 25569) * 86400 * 1000));
        const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        return `${date.getDate()} ${namaBulan[date.getMonth()]} ${date.getFullYear()}`;
    }
    return val.toString().trim();
}

// Ubah value tanggal (serial Excel ATAU teks bebas kayak "Tgl 16 jam 12 siang") jadi objek Date asli,
// biar bisa dibandingkan/overlap-check secara kalender, bukan cuma dibandingkan sebagai teks.
// bulanFallback & tahunFallback dipakai kalau teksnya cuma nyebut angka tanggal doang (misal loading
// date yang cuma ditulis "Tgl 16" tanpa bulan) - biasanya diisi dari tanggal event di event yang sama.
const NAMA_BULAN_LOWER = ["januari", "februari", "maret", "april", "mei", "juni", "juli", "agustus", "september", "oktober", "november", "desember"];

function parseTanggalKeDate(val, bulanFallback = null, tahunFallback = null) {
    if (val === undefined || val === null || val === "" || val === "-") return null;

    // Kasus umum: serial date Excel (angka)
    if (!isNaN(val) && Number(val) > 40000) {
        const utc = new Date(Math.round((Number(val) - 25569) * 86400 * 1000));
        return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
    }

    // Kasus teks bebas: coba tebak tanggal/bulan/tahun dari teksnya
    const teks = val.toString().trim();
    const bulanRegex = new RegExp(NAMA_BULAN_LOWER.join("|"), "i");
    const bulanMatch = teks.match(bulanRegex);
    const angkaMatch = teks.match(/\d{1,2}/); // angka pertama yang muncul dianggap tanggal
    if (!angkaMatch) return null;

    const tanggalNum = parseInt(angkaMatch[0]);
    if (tanggalNum < 1 || tanggalNum > 31) return null;

    const bulanNum = bulanMatch ? NAMA_BULAN_LOWER.indexOf(bulanMatch[0].toLowerCase()) + 1 : bulanFallback;
    const tahunMatch = teks.match(/\d{4}/);
    const tahunNum = tahunMatch ? parseInt(tahunMatch[0]) : tahunFallback;

    if (!bulanNum || !tahunNum) return null;
    return new Date(tahunNum, bulanNum - 1, tanggalNum);
}

// Filter noise di kolom CREW - kadang admin nulis catatan di kolom ini
// (misal "Loding tgl 13 Agustus"), bukan nama crew beneran.
function namaCrewValid(teks) {
    if (!teks) return false;
    const upper = teks.toUpperCase();
    if (upper === "CREW" || upper === "-" || upper === "NTRL") return true; // NTRL = kode internal, tetap dianggap valid
    if (/\d/.test(teks)) return false; // ada angka -> kemungkinan besar catatan tanggal/jam
    if (/TGL|LODING|LOADING|\bJAM\b|CATATAN|NOTE/i.test(teks)) return false;
    if (teks.length > 25) return false; // nama crew biasanya pendek, teks panjang = kemungkinan catatan
    return true;
}

function cleanStr(str) {
    if (!str) return "";
    const s = str.toString().trim();
    const upper = s.toUpperCase();
    if (s === "-" || upper === "EVENT TITTLE" || upper === "VENUE" || upper === "COMPANY" || upper === "NAME") return "";
    return s;
}

function tentukanKategori(namaAlat) {
    const teks = namaAlat.toLowerCase();
    if (teks.includes("drum riser") || teks.includes("riser")) return "🏗️ RIGGING & STAGING";
    if (teks.includes("genset lighting") || teks.includes("panel visual") || teks.includes("panel audio")) return "⚡ POWER";
    if (teks.includes("video mixer") || teks.includes("black magic")) return "📺 VISUAL & MULTIMEDIA";
    if (teks.includes("stage i/o") || teks.includes("analog snake")) return "🔊 SOUND & BACKLINE";

    const kamusKategori = {
        "⚡ POWER": ["genset", "kabel", "power", "panel", "distro"],
        "💡 LIGHTING": ["moving", "strobe", "fresnel", "par led", "par light", "nuovoled", "avolite", "grandma", "grand ma", "lighting", "beam", "smoke", "hazer", "efx", "minuit", "tripod t", "follow spot", "folow spot", "spot led", "blinder", "par zoom", "atomic"],
        "🔊 SOUND & BACKLINE": ["console", "speaker", "subwoofer", "mic", "yamaha", "midas", "dl32", "foh", "mixer", "in ear", "stand mic", "audio focus", "iem", "drumset", "tama", "sound system", "milan", "sp milan", "pa ", "senheiser", "sennheiser", "roland", "akustika", "stage monitor", "musician monitor", "dbr", "dxs", "audio", "pdp", "dw", "cymbal", "paiste", "amplifier", "gallien", "krueger", "head", "snake"],
        "📺 VISUAL & MULTIMEDIA": ["videotron", "tv", "monitor", "projector", "screen", "kamera", "camera", "cam ", "switcher", "klicker", "perfect cue", "laptop", "timer", "sony", "hollyland", "streaming", "vmix", "internet", "orbit", "vj", "visual", "procesor", "processor", "magimage", "led outdoor", "led p", "black magic", "blackmagic"],
        "🏗️ RIGGING & STAGING": ["rigging", "rig", "gawangan", "level", "aluminium", "stage", "barikade", "baricade", "tenda", "mojo"]
    };
    for (const [kategori, kataKunciArray] of Object.entries(kamusKategori)) {
        if (kataKunciArray.some((kataKunci) => teks.includes(kataKunci))) return kategori;
    }
    return "📦 LAINNYA";
}

// ============================================================================
// 4. ENGINE PARSING GOOGLE SHEETS JADWAL
// ============================================================================
// PENTING: fungsi ini SENGAJA tidak menelan error jadi `null` untuk kasus fetch/parse gagal.
// Dulu semua error (network gagal, auth gagal, dsb) ditangkap diam-diam lalu dianggap "[]" (kosong),
// yang artinya kalau bulan ini gagal diambil tapi bulan depan berhasil, hasil gabungannya jadi
// data PARSIAL (cuma bulan depan doang) tapi tetap dianggap "data terbaru yang valid" oleh jalankanRonda.
// Ini yang bikin SEMUA event bulan ini keliatan "hilang" (bug: event dihapus padahal enggak),
// terus muncul lagi sebagai "event baru" begitu fetch berikutnya berhasil normal lagi.
// Sekarang: error fetch/parse asli DILEMPAR ke atas, biar jalankanRonda tau ronda ini gagal total
// dan skip (gak nimpa objekDataLama dengan data yang gak lengkap).
async function getJadwalDariExcel(targetDateObj) {
    const authClient = await authorize();
    if (!authClient) throw new Error("Gagal autentikasi Google Drive (authorize() mengembalikan null)");
    const drive = google.drive({ version: "v3", auth: authClient });

    const res = await drive.files.get({ fileId: SPREADSHEET_ID, alt: "media" }, { responseType: "arraybuffer" });
    const workbook = XLSX.read(res.data, { type: "buffer" });
    const namaBulan = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
    const targetSheetName = `${namaBulan[targetDateObj.getMonth()]} ${targetDateObj.getFullYear()}`;

    const worksheet = workbook.Sheets[targetSheetName];
    // Sheet memang belum dibuat (misal tahun depan) - ini BUKAN error, wajar kosong.
    if (!worksheet) return [];

    return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
}

// Sengaja TIDAK ada try/catch di sini. Kalau salah satu bulan gagal diambil (network error,
// auth expired, dll), biarkan error-nya menjalar ke pemanggilnya (jalankanRonda dkk) supaya
// ronda tsb di-skip total, BUKAN dianggap "data terbaru" yang cuma separuh.
async function getJadwalMultiBulan() {
    const dateIni = new Date();
    const dateDepan = new Date();
    dateDepan.setMonth(dateDepan.getMonth() + 1);

    const dataIni = await getJadwalDariExcel(dateIni);
    const dataDepan = await getJadwalDariExcel(dateDepan);
    return [...dataIni, ...dataDepan];
}

// Helper aman buat dipakai di message handler: pakai cache kalau ada (responsif),
// fallback fetch langsung kalau cache belum ada, dan gak pernah nge-throw ke caller.
async function ambilDataJadwal() {
    if (objekDataLama) return objekDataLama;
    try {
        return await getJadwalMultiBulan();
    } catch (err) {
        console.log("⚠️ Gagal fetch jadwal (fallback ke array kosong):", err.message);
        return [];
    }
}

function prosesDataKePesanWA(rawData, tanggalAngka = "", keywordCari = "") {
    if (!rawData || !Array.isArray(rawData)) return [];
    let daftarPesanWA = [];
    let blocks = [];
    let currentBlock = [];

    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        const colA = row && row[0] ? row[0].toString().trim() : "";
        if (/^\d+$/.test(colA)) {
            if (currentBlock.length > 0) blocks.push(currentBlock);
            currentBlock = [row];
        } else if (currentBlock.length > 0) {
            currentBlock.push(row);
        }
    }
    if (currentBlock.length > 0) blocks.push(currentBlock);

    for (const block of blocks) {
        const masterTanggal = block[0][0].toString().trim();
        if (tanggalAngka !== "" && masterTanggal !== tanggalAngka) continue;

        for (let c = 2; c < 50; c += 8) {
            const getVal = (r, col) => block[r] && block[r][col] ? block[r][col].toString().trim() : "";
            if (getVal(1, c).toUpperCase() !== "NAME") continue;

            const picName = cleanStr(getVal(1, c + 1));
            const companyName = cleanStr(getVal(2, c + 1));
            let eventTitle = cleanStr(getVal(2, c + 6));
            const venue = cleanStr(getVal(3, c + 6));

            if (!picName && !companyName && !eventTitle && !venue) continue;
            if (!eventTitle) eventTitle = venue || companyName || picName || "Event Tanpa Nama";

            const dateEventRaw = getVal(1, c + 6);
            const dateEvent = formatTanggalExcel(dateEventRaw);
            const loadingDate = getVal(4, c + 6) || "-";

            if (keywordCari !== "" && tanggalAngka === "" && keywordCari !== "[SHOW_ALL]") {
                if (!isNaN(dateEventRaw)) {
                    const eventDate = new Date(Math.round((dateEventRaw - 25569) * 86400 * 1000));
                    eventDate.setHours(0,0,0,0);
                    const today = new Date();
                    today.setHours(0,0,0,0);
                    if (eventDate < today) continue; 
                }
            }

            let crewList = [];
            let kategoriAlat = {
                "📺 VISUAL & MULTIMEDIA": [], "💡 LIGHTING": [], "🔊 SOUND & BACKLINE": [],
                "🏗️ RIGGING & STAGING": [], "⚡ POWER": [], "📦 LAINNYA": [],
            };

            for (let i = 8; i < block.length; i++) {
                let rowString = "";
                for (let k = c; k <= c + 7; k++) rowString += getVal(i, k).toUpperCase() + "|";
                if (rowString.includes("STATUS|") || rowString.includes("CUSTOMER DETILS|")) break;

                const crew = getVal(i, c + 6);
                if (crew && crew !== "-" && crew.toUpperCase() !== "CREW") {
                    const upCrew = crew.toUpperCase();
                    if (upCrew !== "DONE" && upCrew !== "CANCEL" && upCrew !== "CANCELLED") {
                        if (!crewList.includes(crew)) crewList.push(crew);
                    }
                }

                const item = getVal(i, c + 1);
                const spec = getVal(i, c + 2);
                const qty = getVal(i, c + 3);
                const freq = getVal(i, c + 5);

                if (qty && item && item.toUpperCase() !== "ITEM") {
                    let namaLengkap = `${item} ${spec}`.trim();
                    let teksAlat = `• ${qty} ${namaLengkap}`;
                    if (freq && freq !== "-") teksAlat += ` (${freq})`;
                    teksAlat = teksAlat.replace(/\s+/g, " ").trim();
                    let namaKategori = tentukanKategori(namaLengkap);
                    if (kategoriAlat[namaKategori]) {
                        kategoriAlat[namaKategori].push(teksAlat);
                    } else {
                        kategoriAlat["📦 LAINNYA"].push(teksAlat);
                    }
                }
            }

            let msg = `━━━━━━━━━━━━━━━━━━━━\n📝 *EVENT DETAIL*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
            msg += `📌 *EVENT* : ${eventTitle}\n`;
            msg += `🏢 *CLIENT* : ${companyName || "-"}\n`;
            msg += `📍 *VENUE* : ${venue || "-"}\n`;
            msg += `📅 *DATE* : ${dateEvent}\n`;
            msg += `🚚 *LOADING*: ${loadingDate}\n\n`;
            msg += `━━━━━━━━━━━━━━━━━━━━\n👥 *CREW*\n`;
            msg += crewList.length > 0 ? crewList.map((cr) => `• ${cr}`).join("\n") : `• (Belum ada crew)`;
            msg += `\n\n`;

            for (const [namaKat, listKat] of Object.entries(kategoriAlat)) {
                if (listKat.length > 0) {
                    msg += `━━━━━━━━━━━━━━━━━━━━\n${namaKat}\n`;
                    msg += listKat.join("\n") + `\n\n`;
                }
            }
            msg = msg.trim() + `\n━━━━━━━━━━━━━━━━━━━━`;

            if (keywordCari !== "" && keywordCari !== "[SHOW_ALL]") {
                if (!msg.toLowerCase().includes(keywordCari.toLowerCase())) continue;
            }

            daftarPesanWA.push(msg);
        }
    }
    return daftarPesanWA;
}


// ============================================================================
// ENGINE FINANCIAL SUITE — DETAIL LENGKAP + PERIODE KONTRAK
// ============================================================================
async function parserEngineSlipGaji(namaTarget, queryBulan = null) {
    const authClient = await authorize();
    if (!authClient) return "❌ Sistem gagal mengakses akun Google.";
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    try {
        const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID_GAJI });
        const allSheets = spreadsheet.data.sheets.map(s => s.properties.title);

        const namaBulan = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
        const dateObj = new Date();
        const blnSekarang = namaBulan[dateObj.getMonth()];
        const thnSekarang = dateObj.getFullYear().toString();

        let sheetAktif = allSheets[allSheets.length - 1];
        if (queryBulan) {
            let found = allSheets.find(s => s.toUpperCase().includes(queryBulan.toUpperCase()));
            if (found) sheetAktif = found;
            else return `ℹ️ Arsip bulan *${queryBulan.toUpperCase()}* tidak ditemukan di Google Sheets.`;
        } else {
            for (let s of allSheets) {
                if (s.toUpperCase().includes(blnSekarang) && s.includes(thnSekarang)) {
                    sheetAktif = s;
                    break;
                }
            }
        }

        // Tentukan window siklus kontrak berdasarkan sheet yang aktif ditampilkan
        const { bulanNum: bulanAktifNum, tahun: tahunAktif } = parseSheetBulanTahun(sheetAktif);
        const periode = getPeriodeKontrak(bulanAktifNum, tahunAktif);

        const batchResponse = await sheets.spreadsheets.values.batchGet({
            spreadsheetId: SPREADSHEET_ID_GAJI,
            ranges: allSheets,
        });

        const allData = batchResponse.data.valueRanges;
        let totalGajiKontrak = 0;
        let totalTabunganKontrak = 0;
        let bulanTerhitungKontrak = 0;
        let dataBulanIni = null;

        allData.forEach((sheetData, index) => {
             const sheetName = allSheets[index];
             const rawData = sheetData.values;
             if (!rawData) return;

             const { bulanNum: sheetBulanNum, tahun: sheetTahun } = parseSheetBulanTahun(sheetName);
             const masukKontrakBerjalan = sheetMasukKontrak(sheetBulanNum, sheetTahun, periode);

             let tCol = -1;
             for (let r = 0; r < Math.min(5, rawData.length); r++) {
                 if (rawData[r]) {
                     for (let c = 0; c < rawData[r].length; c++) {
                         if (rawData[r][c] && rawData[r][c].toString().toUpperCase() === namaTarget.toUpperCase()) {
                             tCol = c; break;
                         }
                     }
                 }
                 if (tCol !== -1) break;
             }
             if (tCol === -1) tCol = 2;

             let sumBulan = 0, tabunganBulan = 0;
             for (let i = 0; i < rawData.length; i++) {
                 const row = rawData[i];
                 if (!row) continue;
                 const c1 = row[1] ? row[1].toString().toUpperCase().trim() : "";
                 const vTarget = row[tCol];
                 let nilai = 0;

                 if (vTarget !== undefined && vTarget !== null && vTarget !== "") {
                     let valStr = vTarget.toString().replace(/\./g, '').replace(/,/g, '').replace(/[^\d-]/g, '');
                     if (!isNaN(valStr) && valStr !== "") nilai = parseInt(valStr);
                 }

                 if (c1 === "SUMMARY") sumBulan = nilai;
                 if (c1 === "TABUNGAN" && (!row[0] || row[0].toString().trim() === "")) tabunganBulan = nilai;
             }

             // Hanya akumulasi kalau sheet ini termasuk siklus kontrak yang sedang berjalan
             if (sumBulan > 0 && masukKontrakBerjalan) {
                 totalGajiKontrak += sumBulan;
                 totalTabunganKontrak += tabunganBulan;
                 bulanTerhitungKontrak++;
             }

             if (sheetName === sheetAktif) dataBulanIni = rawData;
        });

        if (!dataBulanIni) return `ℹ️ Data slip gaji untuk ${sheetAktif} belum tersedia.`;

        const rataRataGaji = bulanTerhitungKontrak > 0 ? totalGajiKontrak / bulanTerhitungKontrak : 0;

        let targetColIndex = -1;
        for (let r = 0; r < Math.min(5, dataBulanIni.length); r++) {
            if (dataBulanIni[r]) {
                for (let c = 0; c < dataBulanIni[r].length; c++) {
                    if (dataBulanIni[r][c] && dataBulanIni[r][c].toString().toUpperCase() === namaTarget.toUpperCase()) {
                        targetColIndex = c; break;
                    }
                }
            }
            if (targetColIndex !== -1) break;
        }
        if (targetColIndex === -1) {
            return `⚠️ Nama "${namaTarget}" tidak ditemukan di header sheet ${sheetAktif}. Cek penulisan nama.`;
        }

        let mode = 0;
        let dataPeriode1 = [];
        let dataPeriode2 = [];
        let vars = {
            totalFee1: 0, kasbon1: 0, transfer1: 0,
            totalFee2: 0, kasbon2: 0, liburLebih2: 0, sisaLibur2: 0, thr2: 0,
            gajiPokok: 0, grandTotal: 0,
            potongTabungan: 0, transfer2: 0, summaryBulan: 0,
        };

        for (let i = 0; i < dataBulanIni.length; i++) {
            const row = dataBulanIni[i];
            if (!row) continue;
            const col0 = row[0] ? row[0].toString().toUpperCase().trim() : "";
            const col1 = row[1] ? row[1].toString().toUpperCase().trim() : "";
            const valTarget = row[targetColIndex];

            let nilai = 0;
            if (valTarget !== undefined && valTarget !== null && valTarget !== "") {
                let valStr = valTarget.toString().replace(/\./g, '').replace(/,/g, '').replace(/[^\d-]/g, '');
                if (!isNaN(valStr) && valStr !== "") nilai = parseInt(valStr);
            }

            if (col0.includes("RECAP FEE TGL 1 -")) { mode = 1; continue; }
            if (col0.includes("RECAP FEE TGL 16") || col0.includes("RECAP FEE TGL 15")) { mode = 2; continue; }

            if (mode === 1) {
                if (col1 === "TOTAL FEE") vars.totalFee1 = nilai;
                if (col1 === "POTONGAN KASBON") vars.kasbon1 = nilai;
                if (col1 === "TOTAL TRANSFER") vars.transfer1 = nilai;
                if (col0 !== "" && col0 !== "TGL" && col1 !== "" && col1 !== "VENUE" && col1 !== "TOTAL FEE" && col1 !== "TOTAL TRANSFER" && nilai > 0) {
                    dataPeriode1.push({ tgl: col0, venue: col1, fee: nilai });
                }
            }

            if (mode === 2) {
                if (col1 === "TOTAL FEE") vars.totalFee2 = nilai;
                if (col1 === "POTONGAN KASBON") vars.kasbon2 = nilai;
                if (col1 === "LIBUR LEBIH") vars.liburLebih2 = nilai;
                if (col1 === "SISA LIBUR") vars.sisaLibur2 = nilai;
                if (col1 === "THR") vars.thr2 = nilai;
                if (col1 === "GAJI POKOK") vars.gajiPokok = nilai;
                if (col1 === "GRAND TOTAL") vars.grandTotal = nilai;
                if (col1 === "TABUNGAN" && col0 === "") vars.potongTabungan = nilai;
                if (col1 === "TOTAL TRANSFER") vars.transfer2 = nilai;
                if (col1 === "SUMMARY") vars.summaryBulan = nilai;

                if (col0 !== "" && col0 !== "TGL" && col0 !== "TABUNGAN" && !col0.includes("BULAN") && col1 !== "" && col1 !== "VENUE" && !col1.includes("TOTAL") && !col1.includes("SUMMARY") && !col1.includes("POTONGAN") && !col1.includes("GAJI") && !col1.includes("GRAND") && col1 !== "TABUNGAN" && nilai > 0) {
                    dataPeriode2.push({ tgl: col0, venue: col1, fee: nilai });
                }
            }
        }

        const sortDateKey = (tglStr) => {
            let clean = tglStr.split('-')[0].trim();
            let num = parseInt(clean);
            return isNaN(num) ? 99 : num;
        };
        dataPeriode1.sort((a, b) => sortDateKey(a.tgl) - sortDateKey(b.tgl));
        dataPeriode2.sort((a, b) => sortDateKey(a.tgl) - sortDateKey(b.tgl));

        const rp = (num) => "Rp " + num.toLocaleString("id-ID");

        const totalFeeJobGabungan = vars.totalFee1 + vars.totalFee2;
        const totalKasbonGabungan = vars.kasbon1 + vars.kasbon2;
        const totalGajiBulanIni = vars.summaryBulan > 0 ? vars.summaryBulan : (vars.transfer1 + vars.transfer2);

        let msg = `━━━━━━━━━━━━━━━━━━━━\n💰 *SLIP GAJI: ${sheetAktif}*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        msg += `👤 *NAMA:* ${namaTarget}\n`;
        msg += `📆 *Periode Kontrak:* ${periode.label}\n`;
        msg += `🔢 *Bulan ke:* ${periode.bulanKe} dari 12\n\n`;

        // 1. RINCIAN JOB
        msg += `📋 *RINCIAN JOB*\n`;
        msg += `_Tgl 1-15:_\n`;
        msg += dataPeriode1.length === 0
            ? `_(Tidak ada job)_\n`
            : dataPeriode1.map(j => `• Tgl ${j.tgl}: ${j.venue} - ${rp(j.fee)}`).join("\n") + "\n";
        msg += `\n_Tgl 16-akhir bulan:_\n`;
        msg += dataPeriode2.length === 0
            ? `_(Tidak ada job)_\n`
            : dataPeriode2.map(j => `• Tgl ${j.tgl}: ${j.venue} - ${rp(j.fee)}`).join("\n") + "\n";
        msg += `━━━━━━━━━━━━━━━━━━━━\n\n`;

        // 2. RINGKASAN GAJI — semua komponen ditampilkan detail
        msg += `📊 *RINGKASAN GAJI BULAN INI*\n`;
        msg += `Total Fee Job (Periode 1+2): ${rp(totalFeeJobGabungan)}\n`;
        msg += totalKasbonGabungan > 0
            ? `Potongan Kasbon: -${rp(totalKasbonGabungan)}\n`
            : `Potongan Kasbon: -\n`;
        msg += vars.liburLebih2 > 0
            ? `Potongan Libur Lebih: -${rp(vars.liburLebih2)}\n`
            : `Potongan Libur Lebih: -\n`;
        msg += vars.sisaLibur2 > 0
            ? `Sisa Libur: +${rp(vars.sisaLibur2)}\n`
            : `Sisa Libur: -\n`;
        msg += vars.thr2 > 0
            ? `THR: +${rp(vars.thr2)}\n`
            : `THR: -\n`;
        msg += `Gaji Pokok: +${rp(vars.gajiPokok)}\n`;
        msg += `— — — — — — — —\n`;
        msg += `*Subtotal (Grand Total):* ${rp(vars.grandTotal)}\n`;
        msg += vars.potongTabungan > 0
            ? `Potong Tabungan: -${rp(vars.potongTabungan)}\n`
            : `Potong Tabungan: -\n`;
        msg += `— — — — — — — —\n`;
        msg += `*💵 TOTAL DITERIMA BULAN INI: ${rp(totalGajiBulanIni)}*\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━\n\n`;

        // 3. JADWAL PENCAIRAN
        msg += `💸 *JADWAL PENCAIRAN*\n`;
        msg += `• Cair Tgl 1-15: ${rp(vars.transfer1)}\n`;
        msg += `• Cair Tgl 16-akhir: ${rp(vars.transfer2)}\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━\n\n`;

        // 4. INFO KONTRAK & TABUNGAN (dihitung per siklus kontrak, bukan tahun kalender)
        msg += `📈 *RINGKASAN SIKLUS KONTRAK (${periode.label})*\n`;
        msg += `• Bulan Tercatat: ${bulanTerhitungKontrak} dari ${periode.bulanKe} bulan berjalan\n`;
        msg += `• Rata-rata Gaji/Bulan: ${rp(Math.round(rataRataGaji))}\n`;
        msg += `• Akumulasi Tabungan: *${rp(totalTabunganKontrak)}*\n`;
        msg += `\n━━━━━━━━━━━━━━━━━━━━`;

        return msg;
    } catch (error) {
        console.error("Error Slip Gaji:", error);
        return `🚨 Error API Slip Gaji: ${error.message}`;
    }
}

async function getRekapTahunan(namaTarget) {
    const authClient = await authorize();
    if (!authClient) return "❌ Sistem gagal mengakses akun Google.";
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    try {
        const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID_GAJI });
        const allSheets = spreadsheet.data.sheets.map(s => s.properties.title);
        const thnSekarang = new Date().getFullYear().toString();

        const batchResponse = await sheets.spreadsheets.values.batchGet({
            spreadsheetId: SPREADSHEET_ID_GAJI,
            ranges: allSheets,
        });

        const allData = batchResponse.data.valueRanges;
        let totalGajiTahunIni = 0;
        let totalTabunganTahunIni = 0;
        let bulanTerhitung = 0;
        let rincianPerBulan = [];

        allData.forEach((sheetData, index) => {
            const sheetName = allSheets[index];
            const rawData = sheetData.values;
            if (!rawData) return;

            let tCol = 2;
            for (let r = 0; r < Math.min(5, rawData.length); r++) {
                if (rawData[r]) {
                    for (let c = 0; c < rawData[r].length; c++) {
                        if (rawData[r][c] && rawData[r][c].toString().toUpperCase() === namaTarget.toUpperCase()) {
                            tCol = c; break;
                        }
                    }
                }
                if (tCol !== 2) break;
            }

            let trans1 = 0, trans2 = 0, tabunganBulan = 0;
            let m = 0;

            for (let i = 0; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row) continue;
                const c0 = row[0] ? row[0].toString().toUpperCase().trim() : "";
                const c1 = row[1] ? row[1].toString().toUpperCase().trim() : "";
                const vTarget = row[tCol];
                let nilai = 0;

                if (vTarget !== undefined && vTarget !== null && vTarget !== "") {
                    let valStr = vTarget.toString().replace(/\./g, '').replace(/,/g, '').replace(/[^\d-]/g, '');
                    if (!isNaN(valStr) && valStr !== "") nilai = parseInt(valStr);
                }

                if (c0.includes("TGL 1 -")) m = 1;
                else if (c0.includes("TGL 15") || c0.includes("TGL 16")) m = 2;

                if (c1 === "TOTAL TRANSFER") {
                    if (m === 1) trans1 = nilai;
                    else if (m === 2) trans2 = nilai;
                }
                if (c1 === "TABUNGAN" && c0 === "") tabunganBulan = nilai;
            }

            let totalBulan = trans1 + trans2;
            if (totalBulan > 0) {
                totalGajiTahunIni += totalBulan;
                totalTabunganTahunIni += tabunganBulan;
                bulanTerhitung++;
                rincianPerBulan.push({ bulan: sheetName, total: totalBulan });
            }
        });

        const rataRataGaji = bulanTerhitung > 0 ? totalGajiTahunIni / bulanTerhitung : 0;
        const rp = (num) => "Rp " + num.toLocaleString("id-ID");

        let msg = `━━━━━━━━━━━━━━━━━━━━\n📈 *REKAP TAHUNAN (${thnSekarang})*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        msg += `👤 *NAMA:* ${namaTarget}\n`;
        msg += `⏱️ *Periode Tercatat:* ${bulanTerhitung} Bulan\n\n`;
        msg += `*📊 RINGKASAN UTAMA:*\n`;
        msg += `• Total Pendapatan (YTD): *${rp(totalGajiTahunIni)}*\n`;
        msg += `• Rata-rata Gaji/Bulan: *${rp(Math.round(rataRataGaji))}*\n`;
        msg += `• Total Akumulasi Tabungan: *${rp(totalTabunganTahunIni)}*\n\n`;
        msg += `*📅 RINCIAN PER BULAN:*\n`;
        rincianPerBulan.forEach(item => { msg += `• ${item.bulan}: ${rp(item.total)}\n`; });
        msg += `\n━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return `🚨 Error Rekap Tahunan: ${error.message}`;
    }
}

async function getTopVenue(namaTarget) {
    const authClient = await authorize();
    if (!authClient) return "❌ Sistem gagal mengakses akun Google.";
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    try {
        const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID_GAJI });
        const allSheets = spreadsheet.data.sheets.map(s => s.properties.title);

        const batchResponse = await sheets.spreadsheets.values.batchGet({
            spreadsheetId: SPREADSHEET_ID_GAJI,
            ranges: allSheets,
        });

        const allData = batchResponse.data.valueRanges;
        let venueMap = {};

        allData.forEach((sheetData) => {
            const rawData = sheetData.values;
            if (!rawData) return;

            let tCol = 2;
            for (let r = 0; r < Math.min(5, rawData.length); r++) {
                if (rawData[r]) {
                    for (let c = 0; c < rawData[r].length; c++) {
                        if (rawData[r][c] && rawData[r][c].toString().toUpperCase() === namaTarget.toUpperCase()) {
                            tCol = c; break;
                        }
                    }
                }
                if (tCol !== 2) break;
            }

            for (let i = 0; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row) continue;
                const col0 = row[0] ? row[0].toString().toUpperCase().trim() : "";
                const col1 = row[1] ? row[1].toString().trim() : "";
                const valTarget = row[tCol];

                let nilai = 0;
                if (valTarget !== undefined && valTarget !== null && valTarget !== "") {
                    let valStr = valTarget.toString().replace(/\./g, '').replace(/,/g, '').replace(/[^\d-]/g, '');
                    if (!isNaN(valStr) && valStr !== "") nilai = parseInt(valStr);
                }

                if (col0 !== "" && col0 !== "TGL" && col0 !== "TABUNGAN" && !col0.includes("BULAN") && col1 !== "" && col1.toUpperCase() !== "VENUE" && !col1.toUpperCase().includes("TOTAL") && !col1.toUpperCase().includes("SUMMARY")) {
                    if (nilai > 0) {
                        let venueName = col1;
                        if (!venueMap[venueName]) venueMap[venueName] = { count: 0, totalFee: 0 };
                        venueMap[venueName].count += 1;
                        venueMap[venueName].totalFee += nilai;
                    }
                }
            }
        });

        let sortedVenues = Object.keys(venueMap).map(v => ({
            name: v,
            count: venueMap[v].count,
            totalFee: venueMap[v].totalFee
        })).sort((a, b) => b.totalFee - a.totalFee).slice(0, 5);

        const rp = (num) => "Rp " + num.toLocaleString("id-ID");
        let msg = `━━━━━━━━━━━━━━━━━━━━\n🏆 *TOP 5 VENUE / KLIEN TERBESAR*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        msg += `👤 *NAMA:* ${namaTarget}\n\n`;

        sortedVenues.forEach((v, idx) => {
            msg += `${idx + 1}. *${v.name}*\n`;
            msg += `   • Total Fee: ${rp(v.totalFee)}\n`;
            msg += `   • Frekuensi: ${v.count} Event\n\n`;
        });
        msg += `━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return `🚨 Error Top Venue: ${error.message}`;
    }
}

async function getProyeksiTabungan(namaTarget) {
    const authClient = await authorize();
    if (!authClient) return "❌ Sistem gagal mengakses akun Google.";
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    try {
        const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID_GAJI });
        const allSheets = spreadsheet.data.sheets.map(s => s.properties.title);

        const batchResponse = await sheets.spreadsheets.values.batchGet({
            spreadsheetId: SPREADSHEET_ID_GAJI,
            ranges: allSheets,
        });

        const allData = batchResponse.data.valueRanges;
        let totalTabunganSekarang = 0;
        let bulanAktifCount = 0;

        allData.forEach((sheetData) => {
            const rawData = sheetData.values;
            if (!rawData) return;
            let tCol = 2;
            for (let r = 0; r < Math.min(5, rawData.length); r++) {
                if (rawData[r]) {
                    for (let c = 0; c < rawData[r].length; c++) {
                        if (rawData[r][c] && rawData[r][c].toString().toUpperCase() === namaTarget.toUpperCase()) {
                            tCol = c; break;
                        }
                    }
                }
                if (tCol !== 2) break;
            }

            let tabunganBulan = 0;
            let adaGaji = false;
            for (let i = 0; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row) continue;
                const c0 = row[0] ? row[0].toString().toUpperCase().trim() : "";
                const c1 = row[1] ? row[1].toString().toUpperCase().trim() : "";
                const vTarget = row[tCol];
                let nilai = 0;

                if (vTarget !== undefined && vTarget !== null && vTarget !== "") {
                    let valStr = vTarget.toString().replace(/\./g, '').replace(/,/g, '').replace(/[^\d-]/g, '');
                    if (!isNaN(valStr) && valStr !== "") nilai = parseInt(valStr);
                }

                if (c1 === "TOTAL TRANSFER" && nilai > 0) adaGaji = true;
                if (c1 === "TABUNGAN" && c0 === "") tabunganBulan = nilai;
            }
            if (adaGaji) {
                totalTabunganSekarang += tabunganBulan;
                bulanAktifCount++;
            }
        });

        // Proyeksi sisa bulan sampai Desember
        let sisaBulan = Math.max(0, 12 - bulanAktifCount);
        let estimasiRataTabunganBulan = bulanAktifCount > 0 ? (totalTabunganSekarang / bulanAktifCount) : 300000;
        let proyeksiTambahan = sisaBulan * estimasiRataTabunganBulan;
        let prediksiAkhirTahun = totalTabunganSekarang + proyeksiTambahan;

        const rp = (num) => "Rp " + num.toLocaleString("id-ID");
        let msg = `━━━━━━━━━━━━━━━━━━━━\n🐷 *PROYEKSI TABUNGAN TAHUNAN*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        msg += `👤 *NAMA:* ${namaTarget}\n\n`;
        msg += `• Saldo Tabungan Saat Ini: *${rp(totalTabunganSekarang)}*\n`;
        msg += `• Bulan Tercatat: ${bulanAktifCount} Bulan\n`;
        msg += `• Estimasi Sisa Tahun Ini: ${sisaBulan} Bulan\n\n`;
        msg += `🚀 *PREDIKSI SALDO AKHIR DESEMBER:*\n`;
        msg += `👉 *${rp(prediksiAkhirTahun)}*\n\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return `🚨 Error Proyeksi Tabungan: ${error.message}`;
    }
}

// ============================================================================
// ENGINE PAYROLL SELURUH KARYAWAN (spreadsheet TERPISAH dari gaji pribadi JAYAK)
// ============================================================================
// Format nama sheet di spreadsheet ini: "OKTOBER26" (nempel, 2 digit tahun),
// BEDA dengan "AGUSTUS 2026" di spreadsheet gaji pribadi.

const formatRupiah = (num) => "Rp " + (num || 0).toLocaleString("id-ID");

// Ubah teks sel ("262.000", "-", "") jadi angka. "-" dan kosong = 0.
function parseAngkaSheet(v) {
    if (v === undefined || v === null || v === "") return 0;
    const s = v.toString().replace(/\./g, "").replace(/,/g, "").replace(/[^\d-]/g, "");
    if (s === "" || isNaN(s)) return 0;
    return parseInt(s);
}

// Header yang BUKAN nama karyawan (kolom total, catatan, dsb) - jangan dihitung sebagai orang.
const NAMA_BUKAN_KARYAWAN = /^(TOTAL|VENUE|TGL|JUMLAH|SUBTOTAL|BULAN|KETERANGAN|KET|NOTE|CATATAN|DONE)$/i;

// Jarak edit (Levenshtein) buat toleransi typo nama: "tomi" -> "TOMY"
function jarakEdit(a, b) {
    const m = a.length, n = b.length;
    const dp = Array.from({ length: m + 1 }, (_, i) => [i, ...Array(n).fill(0)]);
    for (let j = 1; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
        }
    }
    return dp[m][n];
}

function getNamaSheetPayrollAll(dateObj) {
    const namaBulan = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
    const tahun2Digit = dateObj.getFullYear().toString().slice(-2);
    return `${namaBulan[dateObj.getMonth()]}${tahun2Digit}`;
}

// Terjemahkan error Google API jadi pesan yang jelas (dulu semua error keliatan sama-sama "gagal")
function pesanErrorSheets(error) {
    const kode = error && (error.code || (error.response && error.response.status));
    const teks = (error && error.message) || "";
    if (kode === 403) return "🚫 Akun Google yang dipakai bot belum punya akses ke spreadsheet payroll. Share spreadsheet-nya ke akun tsb (minimal Viewer).";
    if (kode === 404) return "❓ Spreadsheet payroll tidak ditemukan. Cek nilai SPREADSHEET_ID_GAJI_ALL di kode.";
    if (/not supported for this document/i.test(teks)) return "⚠️ File payroll bukan Google Sheets asli (kemungkinan file Excel yang diupload). Convert dulu ke Google Sheets.";
    return `🚨 Error payroll: ${teks}`;
}

// Ambil raw rows 1 sheet bulan dari spreadsheet payroll-all.
// queryBulan opsional ("agustus", atau nama sheet persis "AGUSTUS26"); default = bulan berjalan,
// dan kalau sheet bulan berjalan belum dibuat, mundur ke bulan terakhir yang ada.
async function ambilSheetPayrollAll(queryBulan = null) {
    const authClient = await authorize();
    if (!authClient) throw new Error("Gagal autentikasi Google Drive");
    const sheets = google.sheets({ version: "v4", auth: authClient });

    const spreadsheet = await sheets.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID_GAJI_ALL,
        fields: "sheets.properties.title",
    });
    const allSheets = spreadsheet.data.sheets.map((s) => s.properties.title);

    let sheetAktif = null;
    if (queryBulan) {
        const q = queryBulan.toUpperCase().trim();
        sheetAktif = allSheets.find((s) => s.toUpperCase() === q) || allSheets.find((s) => s.toUpperCase().includes(q));
        if (!sheetAktif) return { error: `ℹ️ Arsip bulan *${q}* tidak ditemukan.\nSheet yang ada: ${allSheets.join(", ")}` };
    } else {
        const sekarang = new Date();
        for (let i = 0; i < 12 && !sheetAktif; i++) {
            const nama = getNamaSheetPayrollAll(new Date(sekarang.getFullYear(), sekarang.getMonth() - i, 1));
            if (allSheets.includes(nama)) sheetAktif = nama;
        }
        if (!sheetAktif) sheetAktif = allSheets[allSheets.length - 1];
    }

    // Nama sheet WAJIB dikutip: "MEI26" kebaca sebagai alamat sel (kolom MEI, baris 26) kalau polos.
    const range = `'${sheetAktif.replace(/'/g, "''")}'`;
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID_GAJI_ALL, range });
    return { sheetAktif, rawData: res.data.values || [] };
}

// Cari SEMUA karyawan + kolomnya + timnya.
// Baris header = baris yang kolom A "TGL" dan kolom B "VENUE". Label tim (kalau ada) ada 1 baris di atasnya.
function cariSemuaKaryawanDanTim(rawData) {
    let baseHeaderRow = -1;
    for (let r = 0; r < Math.min(6, rawData.length); r++) {
        const row = rawData[r];
        if (row && row[0] && row[0].toString().toUpperCase().trim() === "TGL" && row[1] && row[1].toString().toUpperCase().trim() === "VENUE") {
            baseHeaderRow = r;
            break;
        }
    }
    if (baseHeaderRow === -1) return [];

    const rowTim = baseHeaderRow > 0 ? (rawData[baseHeaderRow - 1] || []) : [];
    const rowNama = rawData[baseHeaderRow] || [];

    let karyawanList = [];
    let timSekarang = "LAINNYA";
    for (let c = 2; c < rowNama.length; c++) {
        const timLabel = rowTim[c] ? rowTim[c].toString().trim() : "";
        if (timLabel) timSekarang = timLabel;
        const nama = rowNama[c] ? rowNama[c].toString().trim() : "";
        if (!nama || NAMA_BUKAN_KARYAWAN.test(nama) || /^[\d.,\s-]+$/.test(nama)) continue;
        karyawanList.push({ nama, col: c, tim: timSekarang });
    }
    return karyawanList;
}

// Cocokkan nama ketikan user ke daftar karyawan: persis > diawali > mengandung > typo mirip.
// Hasil: { karyawan, saran, dikoreksi }
function cariKaryawanByNama(karyawanList, query) {
    const q = query.toUpperCase().trim();
    const up = (k) => k.nama.toUpperCase();

    let kandidat = karyawanList.filter((k) => up(k) === q);
    if (kandidat.length >= 1) return { karyawan: kandidat[0], saran: [] };

    kandidat = karyawanList.filter((k) => up(k).startsWith(q));
    if (kandidat.length === 1) return { karyawan: kandidat[0], saran: [] };
    if (kandidat.length > 1) return { karyawan: null, saran: kandidat.map((k) => k.nama) };

    kandidat = karyawanList.filter((k) => up(k).includes(q));
    if (kandidat.length === 1) return { karyawan: kandidat[0], saran: [] };
    if (kandidat.length > 1) return { karyawan: null, saran: kandidat.map((k) => k.nama) };

    const dekat = karyawanList.map((k) => ({ k, d: jarakEdit(up(k), q) })).filter((x) => x.d <= 2).sort((a, b) => a.d - b.d);
    if (dekat.length === 0) return { karyawan: null, saran: [] };
    if (dekat[0].d <= 1 && (dekat.length === 1 || dekat[1].d > dekat[0].d)) return { karyawan: dekat[0].k, saran: [], dikoreksi: true };
    return { karyawan: null, saran: dekat.slice(0, 5).map((x) => x.k.nama) };
}

function pesanKaryawanTidakKetemu(query, sheetAktif, saran) {
    let msg = `⚠️ Karyawan "${query}" tidak ditemukan di sheet ${sheetAktif}.`;
    msg += saran.length > 0 ? `\n\nMaksud kamu: *${saran.join(", ")}*?` : `\n\nKetik \`daftar karyawan\` untuk lihat nama yang terdaftar.`;
    return msg;
}

// Ekstrak ringkasan gaji 1 kolom (= 1 karyawan) dari rawData 1 sheet bulan.
// Logikanya sama dengan parserEngineSlipGaji, diparameterin by kolom biar bisa dipakai buat SEMUA karyawan.
function extractRingkasanKaryawan(rawData, colIndex) {
    let mode = 0;
    let dataPeriode1 = [];
    let dataPeriode2 = [];
    let vars = {
        totalFee1: 0, kasbon1: 0, transfer1: 0,
        totalFee2: 0, kasbon2: 0, liburLebih2: 0, sisaLibur2: 0, sisa2: 0, thr2: 0,
        gajiPokok: 0, grandTotal: 0, potongTabungan: 0, transfer2: 0, summaryBulan: 0,
    };

    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row) continue;
        const col0 = row[0] ? row[0].toString().toUpperCase().trim() : "";
        const col1 = row[1] ? row[1].toString().toUpperCase().trim() : "";
        const nilai = parseAngkaSheet(row[colIndex]);

        if (col0.includes("RECAP FEE TGL 1 -")) { mode = 1; continue; }
        if (col0.includes("RECAP FEE TGL 16") || col0.includes("RECAP FEE TGL 15")) { mode = 2; continue; }

        if (mode === 1) {
            if (col1 === "TOTAL FEE") vars.totalFee1 = nilai;
            if (col1 === "POTONGAN KASBON") vars.kasbon1 = nilai;
            if (col1 === "TOTAL TRANSFER") vars.transfer1 = nilai;
            if (col0 !== "" && col0 !== "TGL" && col1 !== "" && col1 !== "VENUE" && col1 !== "TOTAL FEE" && col1 !== "TOTAL TRANSFER" && nilai > 0) {
                dataPeriode1.push({ tgl: row[0].toString().trim(), venue: row[1] ? row[1].toString().trim() : "", fee: nilai });
            }
        }

        if (mode === 2) {
            if (col1 === "TOTAL FEE") vars.totalFee2 = nilai;
            if (col1 === "POTONGAN KASBON") vars.kasbon2 = nilai;
            if (col1 === "LIBUR LEBIH") vars.liburLebih2 = nilai;
            if (col1 === "SISA LIBUR") vars.sisaLibur2 = nilai;
            if (col1 === "SISA") vars.sisa2 = nilai;
            if (col1 === "THR") vars.thr2 = nilai;
            if (col1 === "GAJI POKOK") vars.gajiPokok = nilai;
            if (col1 === "GRAND TOTAL") vars.grandTotal = nilai;
            if (col1 === "TABUNGAN" && col0 === "") vars.potongTabungan = nilai;
            if (col1 === "TOTAL TRANSFER") vars.transfer2 = nilai;
            if (col1 === "SUMMARY") vars.summaryBulan = nilai;

            if (col0 !== "" && col0 !== "TGL" && col0 !== "TABUNGAN" && !col0.includes("BULAN") && col1 !== "" && col1 !== "VENUE" && !col1.includes("TOTAL") && !col1.includes("SUMMARY") && !col1.includes("POTONGAN") && !col1.includes("GAJI") && !col1.includes("GRAND") && col1 !== "TABUNGAN" && nilai > 0) {
                dataPeriode2.push({ tgl: row[0].toString().trim(), venue: row[1] ? row[1].toString().trim() : "", fee: nilai });
            }
        }
    }

    return { ...vars, dataPeriode1, dataPeriode2 };
}

// Cari penanda "sudah ditransfer" (misal tulisan "Done") di baris TOTAL TRANSFER tiap periode.
// Di sheet kamu: baris TOTAL TRANSFER periode 1 punya "Done" di kolom kanan setelah nama-nama karyawan.
function cariStatusPeriode(rawData) {
    let mode = 0;
    const status = { p1: "", p2: "" };
    for (const row of rawData) {
        if (!row) continue;
        const col0 = row[0] ? row[0].toString().toUpperCase().trim() : "";
        const col1 = row[1] ? row[1].toString().toUpperCase().trim() : "";
        if (col0.includes("RECAP FEE TGL 1 -")) { mode = 1; continue; }
        if (col0.includes("RECAP FEE TGL 16") || col0.includes("RECAP FEE TGL 15")) { mode = 2; continue; }
        if (col1 === "TOTAL TRANSFER" && (mode === 1 || mode === 2)) {
            for (let c = 2; c < row.length; c++) {
                const v = row[c] ? row[c].toString().trim() : "";
                if (v && PENANDA_GAJI_CAIR.test(v)) { status[mode === 1 ? "p1" : "p2"] = v.toUpperCase(); break; }
            }
        }
    }
    return status;
}

const urutTgl = (a, b) => {
    const na = parseInt(a.tgl.split("-")[0]);
    const nb = parseInt(b.tgl.split("-")[0]);
    return (isNaN(na) ? 99 : na) - (isNaN(nb) ? 99 : nb);
};

// FITUR: Slip gaji untuk SIAPA SAJA, rinci per periode (angka-angkanya nyambung dengan sheet)
async function getSlipGajiKaryawan(namaKaryawan, queryBulan = null) {
    try {
        const hasil = await ambilSheetPayrollAll(queryBulan);
        if (hasil.error) return hasil.error;
        const { sheetAktif, rawData } = hasil;

        const karyawanList = cariSemuaKaryawanDanTim(rawData);
        if (karyawanList.length === 0) return `⚠️ Daftar karyawan tidak terbaca di sheet *${sheetAktif}* (baris header "TGL | VENUE | nama..." tidak ditemukan).`;

        const cari = cariKaryawanByNama(karyawanList, namaKaryawan);
        if (!cari.karyawan) return pesanKaryawanTidakKetemu(namaKaryawan, sheetAktif, cari.saran);

        const k = cari.karyawan;
        const r = extractRingkasanKaryawan(rawData, k.col);
        const status = cariStatusPeriode(rawData);
        const rp = formatRupiah;
        const garis = `━━━━━━━━━━━━━━━━━━━━`;
        const baris = (label, nilai, tanda) => nilai > 0 ? `${label}: ${tanda}${rp(nilai)}` : `${label}: -`;

        let msg = `${garis}\n💰 *SLIP GAJI: ${sheetAktif}*\n${garis}\n\n`;
        msg += `👤 *NAMA:* ${k.nama.toUpperCase()}\n`;
        if (k.tim && k.tim !== "LAINNYA") msg += `🏷️ *Tim:* ${k.tim}\n`;
        if (cari.dikoreksi) msg += `🔎 _Dicocokkan dengan nama "${k.nama}"_\n`;
        msg += `\n`;

        // PERIODE 1
        msg += `📋 *PERIODE 1 (Tgl 1-15)*${status.p1 ? `  ✅ ${status.p1}` : ""}\n`;
        msg += r.dataPeriode1.length === 0 ? `_(Tidak ada job)_\n` : [...r.dataPeriode1].sort(urutTgl).map((j) => `• Tgl ${j.tgl}: ${j.venue} - ${rp(j.fee)}`).join("\n") + "\n";
        msg += `${baris("Total Fee", r.totalFee1, "")}\n`;
        if (r.kasbon1 > 0) msg += `${baris("Potongan Kasbon", r.kasbon1, "-")}\n`;
        msg += `*💸 Transfer Periode 1: ${rp(r.transfer1)}*\n`;
        msg += `${garis}\n\n`;

        // PERIODE 2
        msg += `📋 *PERIODE 2 (Tgl 16-akhir bulan)*${status.p2 ? `  ✅ ${status.p2}` : ""}\n`;
        msg += r.dataPeriode2.length === 0 ? `_(Tidak ada job)_\n` : [...r.dataPeriode2].sort(urutTgl).map((j) => `• Tgl ${j.tgl}: ${j.venue} - ${rp(j.fee)}`).join("\n") + "\n";
        msg += `${baris("Total Fee", r.totalFee2, "")}\n`;
        msg += `${baris("Potongan Kasbon", r.kasbon2, "-")}\n`;
        msg += `${baris("Potongan Libur Lebih", r.liburLebih2, "-")}\n`;
        msg += `${baris("Sisa Fee Bersih", r.sisa2, "")}\n`;
        msg += `${baris("Sisa Libur", r.sisaLibur2, "+")}\n`;
        if (r.thr2 > 0) msg += `THR: +${rp(r.thr2)}\n`;
        msg += `${baris("Gaji Pokok", r.gajiPokok, "+")}\n`;
        msg += `— — — — — — — —\n`;
        msg += `*Grand Total:* ${rp(r.grandTotal)}\n`;
        msg += `${baris("Potong Tabungan", r.potongTabungan, "-")}\n`;
        msg += `*💸 Transfer Periode 2: ${rp(r.transfer2)}*\n`;
        msg += `${garis}\n\n`;

        msg += `*💵 TOTAL DITERIMA BULAN INI: ${rp(r.transfer1 + r.transfer2)}*\n`;
        msg += garis;
        return msg;
    } catch (error) {
        console.error("Error Slip Gaji Karyawan:", error);
        return pesanErrorSheets(error);
    }
}

// FITUR: Daftar nama karyawan (per tim) - biar tau penulisan nama yang benar
async function getDaftarKaryawan(queryBulan = null) {
    try {
        const hasil = await ambilSheetPayrollAll(queryBulan);
        if (hasil.error) return hasil.error;
        const { sheetAktif, rawData } = hasil;
        const karyawanList = cariSemuaKaryawanDanTim(rawData);
        if (karyawanList.length === 0) return `⚠️ Daftar karyawan tidak terbaca di sheet ${sheetAktif}.`;

        const perTim = {};
        for (const k of karyawanList) {
            if (!perTim[k.tim]) perTim[k.tim] = [];
            perTim[k.tim].push(k.nama);
        }
        let msg = `━━━━━━━━━━━━━━━━━━━━\n👥 *DAFTAR KARYAWAN: ${sheetAktif}*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        for (const [tim, daftar] of Object.entries(perTim)) {
            if (Object.keys(perTim).length > 1 || tim !== "LAINNYA") msg += `🏷️ *${tim}* (${daftar.length})\n`;
            msg += daftar.map((n) => `• ${n}`).join("\n") + "\n\n";
        }
        msg += `Total: *${karyawanList.length} karyawan*\n━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return pesanErrorSheets(error);
    }
}

// FITUR: Rekap payroll company-wide (total cash-out + per periode + per tim)
async function getRekapPayrollBulanIni(queryBulan = null) {
    try {
        const hasil = await ambilSheetPayrollAll(queryBulan);
        if (hasil.error) return hasil.error;
        const { sheetAktif, rawData } = hasil;

        const karyawanList = cariSemuaKaryawanDanTim(rawData);
        if (karyawanList.length === 0) return `⚠️ Gagal membaca daftar karyawan dari sheet ${sheetAktif}.`;
        const status = cariStatusPeriode(rawData);
        const rp = formatRupiah;

        let total1 = 0, total2 = 0, totalKasbon = 0, totalTabungan = 0, jumlahAktif = 0;
        let perTim = {};
        for (const k of karyawanList) {
            const r = extractRingkasanKaryawan(rawData, k.col);
            const totalOrang = r.transfer1 + r.transfer2;
            const kasbonOrang = r.kasbon1 + r.kasbon2;
            if (totalOrang <= 0 && kasbonOrang <= 0) continue;
            total1 += r.transfer1;
            total2 += r.transfer2;
            totalKasbon += kasbonOrang;
            totalTabungan += r.potongTabungan;
            jumlahAktif++;
            perTim[k.tim] = (perTim[k.tim] || 0) + totalOrang;
        }

        let msg = `━━━━━━━━━━━━━━━━━━━━\n🏢 *REKAP PAYROLL: ${sheetAktif}*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        msg += `💵 *Total Payroll Bulan Ini:* ${rp(total1 + total2)}\n`;
        msg += `   • Periode 1 (Tgl 1-15): ${rp(total1)}${status.p1 ? ` ✅ ${status.p1}` : ""}\n`;
        msg += `   • Periode 2 (Tgl 16-akhir): ${rp(total2)}${status.p2 ? ` ✅ ${status.p2}` : ""}\n`;
        msg += `📉 *Total Kasbon Dipotong:* ${rp(totalKasbon)}\n`;
        msg += `🐷 *Total Tabungan Dipotong:* ${rp(totalTabungan)}\n`;
        msg += `👥 *Jumlah Karyawan Aktif:* ${jumlahAktif}\n`;
        if (Object.keys(perTim).length > 1 || !perTim["LAINNYA"]) {
            msg += `\n📊 *RINCIAN PER TIM*\n`;
            for (const [tim, total] of Object.entries(perTim)) msg += `• ${tim}: ${rp(total)}\n`;
        }
        msg += `\n━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return pesanErrorSheets(error);
    }
}

// FITUR: Daftar kasbon aktif, diurut dari terbesar
async function getDaftarKasbonAktif(queryBulan = null) {
    try {
        const hasil = await ambilSheetPayrollAll(queryBulan);
        if (hasil.error) return hasil.error;
        const { sheetAktif, rawData } = hasil;

        const karyawanList = cariSemuaKaryawanDanTim(rawData);
        const rp = formatRupiah;

        let daftarKasbon = [];
        for (const k of karyawanList) {
            const r = extractRingkasanKaryawan(rawData, k.col);
            const totalKasbon = r.kasbon1 + r.kasbon2;
            if (totalKasbon > 0) daftarKasbon.push({ nama: k.nama, kasbon: totalKasbon, k1: r.kasbon1, k2: r.kasbon2 });
        }
        daftarKasbon.sort((a, b) => b.kasbon - a.kasbon);

        if (daftarKasbon.length === 0) return `✅ Tidak ada kasbon aktif bulan ini (${sheetAktif}).`;

        let msg = `━━━━━━━━━━━━━━━━━━━━\n📉 *DAFTAR KASBON AKTIF: ${sheetAktif}*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        daftarKasbon.forEach((k, idx) => {
            const rinci = k.k1 > 0 && k.k2 > 0 ? ` _(P1 ${rp(k.k1)} + P2 ${rp(k.k2)})_` : "";
            msg += `${idx + 1}. *${k.nama}* - ${rp(k.kasbon)}${rinci}\n`;
        });
        msg += `\n👥 ${daftarKasbon.length} karyawan | 💰 *Total Kasbon:* ${rp(daftarKasbon.reduce((s, k) => s + k.kasbon, 0))}`;
        msg += `\n━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return pesanErrorSheets(error);
    }
}

// FITUR: Leaderboard TOP 10 - fee tertinggi & job terbanyak
async function getLeaderboardPayroll(queryBulan = null) {
    try {
        const hasil = await ambilSheetPayrollAll(queryBulan);
        if (hasil.error) return hasil.error;
        const { sheetAktif, rawData } = hasil;

        const karyawanList = cariSemuaKaryawanDanTim(rawData);
        const rp = formatRupiah;

        let rincian = [];
        for (const k of karyawanList) {
            const r = extractRingkasanKaryawan(rawData, k.col);
            const totalJob = r.dataPeriode1.length + r.dataPeriode2.length;
            const totalFee = r.totalFee1 + r.totalFee2;
            if (totalFee <= 0 && totalJob <= 0) continue;
            rincian.push({ nama: k.nama, totalJob, totalFee });
        }
        if (rincian.length === 0) return `ℹ️ Belum ada data fee/job di sheet ${sheetAktif}.`;

        const medali = (idx) => (idx === 0 ? "🥇" : idx === 1 ? "🥈" : idx === 2 ? "🥉" : `${idx + 1}.`);
        const topFee = [...rincian].sort((a, b) => b.totalFee - a.totalFee).slice(0, 10);
        const topJob = [...rincian].sort((a, b) => b.totalJob - a.totalJob || b.totalFee - a.totalFee).slice(0, 10);

        let msg = `━━━━━━━━━━━━━━━━━━━━\n🏆 *LEADERBOARD TOP 10: ${sheetAktif}*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        msg += `💰 *Top 10 Fee Tertinggi:*\n`;
        topFee.forEach((k, idx) => { msg += `${medali(idx)} ${k.nama} - ${rp(k.totalFee)}\n`; });
        msg += `\n📅 *Top 10 Job Terbanyak:*\n`;
        topJob.forEach((k, idx) => { msg += `${medali(idx)} ${k.nama} - ${k.totalJob} job\n`; });
        msg += `\n_Dari ${rincian.length} karyawan yang punya fee bulan ini._\n━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return pesanErrorSheets(error);
    }
}

// Baca bagian TABUNGAN di bawah sheet: baris "TABUNGAN" -> header "BULAN | nama..." -> baris per bulan -> "TOTAL".
function parseTabunganSection(rawData) {
    let idxHeader = -1;
    for (let i = 0; i < rawData.length - 1; i++) {
        const row = rawData[i];
        if (!row) continue;
        const c0 = row[0] ? row[0].toString().toUpperCase().trim() : "";
        if (c0 === "TABUNGAN") {
            const next = rawData[i + 1];
            if (next && next[1] && next[1].toString().toUpperCase().trim() === "BULAN") { idxHeader = i + 1; break; }
        }
    }
    if (idxHeader === -1) return null;

    const header = rawData[idxHeader];
    const kolom = [];
    for (let c = 2; c < header.length; c++) {
        const n = header[c] ? header[c].toString().trim() : "";
        if (n && !NAMA_BUKAN_KARYAWAN.test(n)) kolom.push({ nama: n, col: c });
    }

    const bulanRows = [];
    let totalRow = null;
    for (let i = idxHeader + 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row) continue;
        const label = row[1] ? row[1].toString().trim() : "";
        if (!label) { if (bulanRows.length > 0) break; else continue; }
        if (label.toUpperCase() === "TOTAL") { totalRow = row; break; }
        bulanRows.push({ label, row });
    }
    return { kolom, bulanRows, totalRow };
}

// FITUR: Tabungan karyawan. namaKaryawan = null -> semua karyawan (ranking total tabungan)
async function getTabunganKaryawan(namaKaryawan = null, queryBulan = null) {
    try {
        const hasil = await ambilSheetPayrollAll(queryBulan);
        if (hasil.error) return hasil.error;
        const { sheetAktif, rawData } = hasil;
        const tab = parseTabunganSection(rawData);
        if (!tab || tab.kolom.length === 0) return `ℹ️ Bagian TABUNGAN tidak ditemukan di sheet ${sheetAktif}.`;
        const rp = formatRupiah;

        const hitungOrang = (k) => {
            let totalHitung = 0;
            const rincian = [];
            for (const b of tab.bulanRows) {
                const n = parseAngkaSheet(b.row[k.col]);
                if (n > 0) { totalHitung += n; rincian.push({ bulan: b.label, nilai: n }); }
            }
            const totalSheet = tab.totalRow ? parseAngkaSheet(tab.totalRow[k.col]) : 0;
            return { nama: k.nama, total: totalSheet > 0 ? totalSheet : totalHitung, rincian };
        };

        if (namaKaryawan) {
            const cari = cariKaryawanByNama(tab.kolom, namaKaryawan);
            if (!cari.karyawan) return pesanKaryawanTidakKetemu(namaKaryawan, sheetAktif, cari.saran);
            const d = hitungOrang(cari.karyawan);
            let msg = `━━━━━━━━━━━━━━━━━━━━\n🐷 *TABUNGAN: ${d.nama.toUpperCase()}*\n━━━━━━━━━━━━━━━━━━━━\n_Sumber: sheet ${sheetAktif}_\n\n`;
            msg += d.rincian.length === 0 ? `_(Belum ada setoran tabungan)_\n` : d.rincian.map((x) => `• ${x.bulan}: ${rp(x.nilai)}`).join("\n") + "\n";
            msg += `\n💰 *Total Tabungan:* ${rp(d.total)}\n━━━━━━━━━━━━━━━━━━━━`;
            return msg;
        }

        const semua = tab.kolom.map(hitungOrang).filter((d) => d.total > 0).sort((a, b) => b.total - a.total);
        if (semua.length === 0) return `ℹ️ Belum ada tabungan tercatat di sheet ${sheetAktif}.`;
        let msg = `━━━━━━━━━━━━━━━━━━━━\n🐷 *TABUNGAN SELURUH KARYAWAN*\n━━━━━━━━━━━━━━━━━━━━\n_Sumber: sheet ${sheetAktif}_\n\n`;
        semua.forEach((d, idx) => { msg += `${idx + 1}. ${d.nama} - ${rp(d.total)}\n`; });
        msg += `\n👥 ${semua.length} karyawan | 💰 *Total Dana Tabungan:* ${rp(semua.reduce((s, d) => s + d.total, 0))}\n━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return pesanErrorSheets(error);
    }
}

// Nama sheet bulan sebelum "OKTOBER26" -> "SEPTEMBER26"
function namaSheetBulanSebelumnya(namaSheet) {
    const namaBulan = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
    const idx = namaBulan.findIndex((n) => namaSheet.toUpperCase().startsWith(n));
    const yy = parseInt(namaSheet.slice(-2));
    if (idx === -1 || isNaN(yy)) return null;
    const prevIdx = idx === 0 ? 11 : idx - 1;
    const prevYY = idx === 0 ? yy - 1 : yy;
    return `${namaBulan[prevIdx]}${String(prevYY).padStart(2, "0")}`;
}

// FITUR: Bandingkan gaji bulan ini vs bulan lalu untuk 1 karyawan
async function getPerbandinganGaji(namaKaryawan) {
    try {
        const skrg = await ambilSheetPayrollAll();
        if (skrg.error) return skrg.error;
        const listSkrg = cariSemuaKaryawanDanTim(skrg.rawData);
        const cari = cariKaryawanByNama(listSkrg, namaKaryawan);
        if (!cari.karyawan) return pesanKaryawanTidakKetemu(namaKaryawan, skrg.sheetAktif, cari.saran);
        const rSkrg = extractRingkasanKaryawan(skrg.rawData, cari.karyawan.col);

        const namaPrev = namaSheetBulanSebelumnya(skrg.sheetAktif);
        if (!namaPrev) return `⚠️ Gagal menentukan bulan sebelum ${skrg.sheetAktif}.`;
        const prev = await ambilSheetPayrollAll(namaPrev);
        if (prev.error) return prev.error;
        const cariPrev = cariKaryawanByNama(cariSemuaKaryawanDanTim(prev.rawData), cari.karyawan.nama);
        if (!cariPrev.karyawan) return `ℹ️ ${cari.karyawan.nama} tidak ada di sheet ${prev.sheetAktif}, jadi belum bisa dibandingkan.`;
        const rPrev = extractRingkasanKaryawan(prev.rawData, cariPrev.karyawan.col);

        const rp = formatRupiah;
        const selisih = (lama, baru) => {
            const d = baru - lama;
            if (d === 0) return "sama";
            const persen = lama > 0 ? ` (${d > 0 ? "+" : ""}${Math.round((d / lama) * 100)}%)` : "";
            return `${d > 0 ? "▲ +" : "▼ -"}${rp(Math.abs(d))}${persen}`;
        };
        const feeS = rSkrg.totalFee1 + rSkrg.totalFee2, feeP = rPrev.totalFee1 + rPrev.totalFee2;
        const jobS = rSkrg.dataPeriode1.length + rSkrg.dataPeriode2.length, jobP = rPrev.dataPeriode1.length + rPrev.dataPeriode2.length;
        const kasS = rSkrg.kasbon1 + rSkrg.kasbon2, kasP = rPrev.kasbon1 + rPrev.kasbon2;
        const trfS = rSkrg.transfer1 + rSkrg.transfer2, trfP = rPrev.transfer1 + rPrev.transfer2;

        let msg = `━━━━━━━━━━━━━━━━━━━━\n📊 *PERBANDINGAN GAJI: ${cari.karyawan.nama.toUpperCase()}*\n━━━━━━━━━━━━━━━━━━━━\n`;
        msg += `${skrg.sheetAktif} vs ${prev.sheetAktif}\n\n`;
        msg += `• Total Fee Job: ${rp(feeP)} ➡️ ${rp(feeS)}\n   ${selisih(feeP, feeS)}\n`;
        msg += `• Jumlah Job: ${jobP} ➡️ ${jobS}\n`;
        msg += `• Potongan Kasbon: ${rp(kasP)} ➡️ ${rp(kasS)}\n`;
        msg += `• Total Diterima: ${rp(trfP)} ➡️ ${rp(trfS)}\n   ${selisih(trfP, trfS)}\n`;
        msg += `\n_Catatan: bulan berjalan belum tentu final._\n━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return pesanErrorSheets(error);
    }
}

// ----------------------------------------------------------------------------
// SNAPSHOT & RONDA PAYROLL (notifikasi otomatis kalau ada slip gaji masuk/berubah)
// ----------------------------------------------------------------------------
function buatSnapshotGaji(sheetAktif, rawData) {
    const karyawan = {};
    for (const k of cariSemuaKaryawanDanTim(rawData)) {
        const r = extractRingkasanKaryawan(rawData, k.col);
        karyawan[k.nama.toUpperCase()] = {
            nama: k.nama,
            tim: k.tim,
            totalFee1: r.totalFee1,
            totalFee2: r.totalFee2,
            transfer1: r.transfer1,
            transfer2: r.transfer2,
            kasbon: r.kasbon1 + r.kasbon2,
            jobs: [
                ...r.dataPeriode1.map((j) => ({ p: 1, tgl: j.tgl, venue: j.venue, fee: j.fee })),
                ...r.dataPeriode2.map((j) => ({ p: 2, tgl: j.tgl, venue: j.venue, fee: j.fee })),
            ],
        };
    }
    return { sheet: sheetAktif, karyawan, status: cariStatusPeriode(rawData) };
}

// Bandingkan snapshot terakhir yang SUDAH diinfokan (lama) dengan kondisi sekarang (baru).
// Hasil: array { tipe: "cair" | "pribadi" | "admin", teks }
function bandingkanSnapshotGaji(lama, baru) {
    const hasil = [];
    const rp = formatRupiah;
    const namaSaya = NAMA_SAYA_DI_ABSEN.toUpperCase();
    const labelPeriode = { p1: "TGL 1-15", p2: "TGL 16-AKHIR BULAN" };

    // 1. Periode baru ditandai sudah ditransfer ("Done")
    for (const kode of ["p1", "p2"]) {
        if (!lama.status[kode] && baru.status[kode]) {
            let total = 0, jumlah = 0;
            for (const k of Object.values(baru.karyawan)) {
                const t = kode === "p1" ? k.transfer1 : k.transfer2;
                if (t > 0) { total += t; jumlah++; }
            }
            let teks = `✅ *GAJI PERIODE ${labelPeriode[kode]} (${baru.sheet}) SUDAH DITRANSFER*\n_Status di sheet: ${baru.status[kode]}_\n`;
            const saya = baru.karyawan[namaSaya];
            if (saya) teks += `\n👤 Gaji kamu (${saya.nama}): *${rp(kode === "p1" ? saya.transfer1 : saya.transfer2)}*`;
            teks += `\n🏢 Total transfer seluruh karyawan: *${rp(total)}* (${jumlah} orang)`;
            hasil.push({ tipe: "cair", teks });
        }
    }

    // 2. Perubahan data gaji pribadi (fee job baru/diubah/dihapus, transfer, kasbon)
    const lamaSaya = lama.karyawan[namaSaya];
    const baruSaya = baru.karyawan[namaSaya];
    if (lamaSaya && baruSaya) {
        const baris = [];
        const petaJob = (jobs) => {
            const peta = {}, hitung = {};
            for (const j of jobs) {
                const dasar = `${j.p}|${j.tgl}|${j.venue}`;
                hitung[dasar] = (hitung[dasar] || 0) + 1;
                peta[`${dasar}#${hitung[dasar]}`] = j;
            }
            return peta;
        };
        const jl = petaJob(lamaSaya.jobs), jb = petaJob(baruSaya.jobs);
        for (const key in jb) {
            if (!jl[key]) baris.push(`➕ Fee baru: Tgl ${jb[key].tgl} - ${jb[key].venue} - ${rp(jb[key].fee)}`);
            else if (jl[key].fee !== jb[key].fee) baris.push(`✏️ Fee diubah: Tgl ${jb[key].tgl} - ${jb[key].venue}: ${rp(jl[key].fee)} ➡️ ${rp(jb[key].fee)}`);
        }
        for (const key in jl) {
            if (!jb[key]) baris.push(`➖ Fee dihapus: Tgl ${jl[key].tgl} - ${jl[key].venue} - ${rp(jl[key].fee)}`);
        }
        if (lamaSaya.transfer1 !== baruSaya.transfer1) baris.push(`💸 Transfer Tgl 1-15: ${rp(lamaSaya.transfer1)} ➡️ ${rp(baruSaya.transfer1)}`);
        if (lamaSaya.transfer2 !== baruSaya.transfer2) baris.push(`💸 Transfer Tgl 16-akhir: ${rp(lamaSaya.transfer2)} ➡️ ${rp(baruSaya.transfer2)}`);
        if (lamaSaya.kasbon !== baruSaya.kasbon) baris.push(`📉 Kasbon: ${rp(lamaSaya.kasbon)} ➡️ ${rp(baruSaya.kasbon)}`);
        if (baris.length > 0) {
            let teks = `💰 *UPDATE DATA GAJI KAMU (${baru.sheet})*\n\n${baris.join("\n")}`;
            teks += `\n\n💵 Total diterima bulan ini saat ini: *${rp(baruSaya.transfer1 + baruSaya.transfer2)}*`;
            hasil.push({ tipe: "pribadi", teks });
        }
    }

    // 3. Ringkasan seluruh karyawan (bisa dimatikan lewat NOTIF_PAYROLL_SEMUA_KARYAWAN)
    if (NOTIF_PAYROLL_SEMUA_KARYAWAN) {
        const hitungTotal = (snap) => Object.values(snap.karyawan).reduce((s, k) => s + k.transfer1 + k.transfer2, 0);
        const totalLama = hitungTotal(lama), totalBaru = hitungTotal(baru);
        let berubah = 0;
        const barisKasbon = [];
        for (const key in baru.karyawan) {
            if (key === namaSaya) continue;
            const b = baru.karyawan[key], l = lama.karyawan[key];
            if (!l || l.transfer1 !== b.transfer1 || l.transfer2 !== b.transfer2 || l.totalFee1 !== b.totalFee1 || l.totalFee2 !== b.totalFee2) berubah++;
            const kasbonLama = l ? l.kasbon : 0;
            if (b.kasbon !== kasbonLama) barisKasbon.push(`• ${b.nama}: ${rp(kasbonLama)} ➡️ ${rp(b.kasbon)}`);
        }
        if (berubah > 0 || barisKasbon.length > 0) {
            let teks = `🏢 *UPDATE PAYROLL SELURUH KARYAWAN (${baru.sheet})*\n`;
            if (berubah > 0) teks += `\n👥 ${berubah} karyawan lain berubah datanya`;
            if (totalLama !== totalBaru) teks += `\n💵 Total payroll: ${rp(totalLama)} ➡️ ${rp(totalBaru)}`;
            if (barisKasbon.length > 0) {
                teks += `\n\n📉 *Perubahan kasbon:*\n${barisKasbon.slice(0, 15).join("\n")}`;
                if (barisKasbon.length > 15) teks += `\n...dan ${barisKasbon.length - 15} lainnya`;
            }
            hasil.push({ tipe: "admin", teks });
        }
    }
    return hasil;
}

async function muatSnapshotGaji() {
    try {
        const isi = await fs.readFile(FILE_SNAPSHOT_GAJI, "utf8");
        const data = JSON.parse(isi);
        if (data && data.sheet && data.karyawan && data.status) snapshotGajiDinotif = data;
    } catch (e) { /* belum ada file snapshot, wajar di run pertama */ }
}

async function simpanSnapshotGaji() {
    try { await fs.writeFile(FILE_SNAPSHOT_GAJI, JSON.stringify(snapshotGajiDinotif)); }
    catch (e) { console.log("⚠️ Gagal simpan snapshot gaji:", e.message); }
}

async function kirimNotifikasiGaji(perubahan) {
    const adminInfo = await getAdminTerakhirEdit(1, SPREADSHEET_ID_GAJI_ALL);
    let header = `🔔 *INFO PAYROLL TERBARU*`;
    if (adminInfo) {
        const waktuFormat = new Date(adminInfo.waktu).toLocaleString("id-ID", { timeZone: "Asia/Makassar", dateStyle: "medium", timeStyle: "short" });
        header += `\n✏️ *Diubah oleh:* ${adminInfo.nama}\n🕒 *Waktu edit:* ${waktuFormat} WITA`;
    } else {
        header += `\n✏️ *Diubah oleh:* (gagal dideteksi)`;
    }
    await kirimWA(ID_TUJUAN_NOTIFIKASI, header);
    for (const p of perubahan) await kirimWA(ID_TUJUAN_NOTIFIKASI, p.teks);

    // Kalau ada periode yang baru cair, kirim sekalian slip gaji terbarunya
    if (perubahan.some((p) => p.tipe === "cair")) {
        const slip = await getSlipGajiKaryawan(NAMA_SAYA_DI_ABSEN);
        await kirimWA(ID_TUJUAN_NOTIFIKASI, `📄 *Slip gaji terbaru kamu:*`);
        await kirimWA(ID_TUJUAN_NOTIFIKASI, slip);
    }
}

const jalankanRondaGaji = async () => {
    if (rondaGajiBerjalan) return;
    rondaGajiBerjalan = true;
    try {
        const hasil = await ambilSheetPayrollAll();
        if (hasil.error) return;
        const snapshotBaru = buatSnapshotGaji(hasil.sheetAktif, hasil.rawData);
        const jumlahBaru = Object.keys(snapshotBaru.karyawan).length;
        if (jumlahBaru === 0) { console.log("⚠️ [Ronda Gaji] daftar karyawan tidak terbaca, ronda di-skip."); return; }
        gagalRondaGajiBeruntun = 0;

        const sekarang = Date.now();

        // Pertama kali / ganti bulan: simpan baseline saja, JANGAN kirim notifikasi
        if (!snapshotGajiDinotif || snapshotGajiDinotif.sheet !== snapshotBaru.sheet) {
            snapshotGajiDinotif = snapshotBaru;
            snapshotGajiTerakhir = snapshotBaru;
            waktuGajiBerubahTerakhir = 0;
            waktuGajiPendingMulai = 0;
            await simpanSnapshotGaji();
            console.log(`💼 [Ronda Gaji] baseline ${snapshotBaru.sheet} tersimpan (${jumlahBaru} karyawan).`);
            return;
        }

        // Sanity check: jumlah karyawan anjlok drastis = data kebaca rusak, bukan beneran dihapus semua
        if (snapshotGajiTerakhir) {
            const jumlahLama = Object.keys(snapshotGajiTerakhir.karyawan).length;
            if (jumlahLama >= 5 && jumlahBaru < jumlahLama * 0.5) {
                console.log(`⚠️ [Ronda Gaji] jumlah karyawan anjlok (${jumlahLama} -> ${jumlahBaru}), ronda di-skip.`);
                return;
            }
            if (JSON.stringify(snapshotBaru) !== JSON.stringify(snapshotGajiTerakhir)) waktuGajiBerubahTerakhir = sekarang;
        }
        snapshotGajiTerakhir = snapshotBaru;

        const perubahan = bandingkanSnapshotGaji(snapshotGajiDinotif, snapshotBaru);
        if (perubahan.length === 0) { waktuGajiPendingMulai = 0; return; }
        if (!waktuGajiPendingMulai) waktuGajiPendingMulai = sekarang;

        // Debounce: tunggu admin selesai ngetik (data tenang JEDA_TENANG_GAJI_MS), kecuali
        // periode baru cair ("Done") yang dikirim langsung, atau sudah menunggu terlalu lama.
        const adaCair = perubahan.some((p) => p.tipe === "cair");
        const tenang = sekarang - waktuGajiBerubahTerakhir >= JEDA_TENANG_GAJI_MS;
        const kelamaan = sekarang - waktuGajiPendingMulai >= BATAS_TUNGGU_GAJI_MS;
        if (adaCair || tenang || kelamaan) {
            snapshotGajiDinotif = snapshotBaru;      // tandai dulu supaya gak kirim dobel kalau pengiriman lambat
            waktuGajiPendingMulai = 0;
            await simpanSnapshotGaji();
            await kirimNotifikasiGaji(perubahan);
        }
    } catch (err) {
        gagalRondaGajiBeruntun++;
        console.log(`❌ [Ronda Gaji] gagal (${gagalRondaGajiBeruntun}x berturut-turut): ${err.message}`);
        if (gagalRondaGajiBeruntun === 6) sendSystemAlert(`❌ Ronda payroll gagal 6x berturut-turut: ${pesanErrorSheets(err)}`);
    } finally {
        rondaGajiBerjalan = false;
    }
};

// Pengingat pencairan: dikirim pagi di tanggal HARI_PENGINGAT_PERIODE_1 (periode 1)
// dan di hari terakhir bulan (periode 2), kecuali periodenya sudah ditandai Done di sheet.
async function kirimPengingatGaji() {
    const sekarang = new Date();
    const hari = sekarang.getDate();
    const akhirBulan = new Date(sekarang.getFullYear(), sekarang.getMonth() + 1, 0).getDate();
    let kode = null;
    if (hari === HARI_PENGINGAT_PERIODE_1) kode = "p1";
    else if (hari === akhirBulan) kode = "p2";
    if (!kode) return;

    const hasil = await ambilSheetPayrollAll();
    if (hasil.error) return;
    const snap = buatSnapshotGaji(hasil.sheetAktif, hasil.rawData);
    if (snap.status[kode]) return;

    let total = 0, jumlah = 0;
    for (const k of Object.values(snap.karyawan)) {
        const t = kode === "p1" ? k.transfer1 : k.transfer2;
        if (t > 0) { total += t; jumlah++; }
    }
    if (jumlah === 0) return;
    const label = kode === "p1" ? "Tgl 1-15" : "Tgl 16-akhir bulan";
    let msg = `⏰ *PENGINGAT PENCAIRAN GAJI*\nHari ini jadwal pencairan periode ${label} (${snap.sheet}).\n\n`;
    msg += `💵 Total yang perlu ditransfer: *${formatRupiah(total)}*\n👥 Jumlah karyawan: ${jumlah}\n\n`;
    msg += `_Ketik \`rekap payroll\` untuk rincian. Pengingat ini berhenti otomatis kalau sheet sudah ditandai Done._`;
    await kirimWA(ID_TUJUAN_NOTIFIKASI, msg);
}

// Status gaji manual (`status gaji`): periode mana yang sudah ditandai cair + nominal
async function getStatusGaji() {
    try {
        const hasil = await ambilSheetPayrollAll();
        if (hasil.error) return hasil.error;
        const snap = buatSnapshotGaji(hasil.sheetAktif, hasil.rawData);
        if (Object.keys(snap.karyawan).length === 0) return `⚠️ Daftar karyawan tidak terbaca di sheet ${snap.sheet}.`;
        const rp = formatRupiah;
        const hitung = (kode) => {
            let total = 0, jumlah = 0;
            for (const k of Object.values(snap.karyawan)) {
                const t = kode === "p1" ? k.transfer1 : k.transfer2;
                if (t > 0) { total += t; jumlah++; }
            }
            return { total, jumlah };
        };
        const p1 = hitung("p1"), p2 = hitung("p2");
        const saya = snap.karyawan[NAMA_SAYA_DI_ABSEN.toUpperCase()];

        let msg = `━━━━━━━━━━━━━━━━━━━━\n📌 *STATUS GAJI: ${snap.sheet}*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
        msg += `*Periode 1 (Tgl 1-15)*\n${snap.status.p1 ? `✅ Sudah ditransfer (${snap.status.p1})` : `⏳ Belum ditandai cair di sheet`}\n💵 ${rp(p1.total)} untuk ${p1.jumlah} karyawan\n\n`;
        msg += `*Periode 2 (Tgl 16-akhir)*\n${snap.status.p2 ? `✅ Sudah ditransfer (${snap.status.p2})` : `⏳ Belum ditandai cair di sheet`}\n💵 ${rp(p2.total)} untuk ${p2.jumlah} karyawan\n`;
        if (saya) msg += `\n👤 *Gaji kamu (${saya.nama}):*\n• Periode 1: ${rp(saya.transfer1)}\n• Periode 2: ${rp(saya.transfer2)}\n• Total: *${rp(saya.transfer1 + saya.transfer2)}*\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━`;
        return msg;
    } catch (error) {
        return pesanErrorSheets(error);
    }
}

// ============================================================================
// 6. ENGINE WHATSAPP & BOT LOGIC
// ============================================================================
const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"] },
});

// Kirim pesan Telegram sebagai fallback, dipakai saat WA down/putus koneksi.
// Pakai modul https bawaan Node biar gak nambah dependency & aman di versi Node manapun.
function kirimTelegramFallback(text) {
    return new Promise((resolve) => {
        if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
            resolve(false);
            return;
        }
        try {
            const payload = JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text, parse_mode: "Markdown" });
            const req = https.request(
                {
                    hostname: "api.telegram.org",
                    path: `/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
                    method: "POST",
                    headers: { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(payload) },
                },
                (res) => {
                    res.on("data", () => {});
                    res.on("end", () => resolve(true));
                }
            );
            req.on("error", (err) => {
                console.log("⚠️ Gagal kirim fallback Telegram:", err.message);
                resolve(false);
            });
            req.write(payload);
            req.end();
        } catch (err) {
            console.log("⚠️ Gagal kirim fallback Telegram:", err.message);
            resolve(false);
        }
    });
}

// Kirim pesan WA lewat antrian serial (rate limit guard), biar gak kena flag spam
// WhatsApp kalau banyak pesan meletus bersamaan (misal banyak revisi terdeteksi sekaligus).
function kirimWA(target, text) {
    antrianKirimWA = antrianKirimWA
        .then(() => client.sendMessage(target, text))
        .catch((err) => console.log("⚠️ Gagal kirim WA:", err.message))
        .then(() => new Promise((res) => setTimeout(res, JEDA_ANTAR_PESAN_MS)));
    return antrianKirimWA;
}

const sendSystemAlert = async (text) => {
    console.log(text);
    if (isBotReady) {
        try {
            await kirimWA(ID_TUJUAN_NOTIFIKASI, text);
            return;
        } catch (e) {}
    }
    // WA belum siap / gagal kirim -> coba fallback ke Telegram
    await kirimTelegramFallback(text);
};

const simulateTyping = async (chat, text) => {
    if (!chat) return;
    try {
        await chat.sendSeen();
        await chat.sendStateTyping();
        let typingTime = Math.min(text.length * 30 + 500, 2000);
        await new Promise((resolve) => setTimeout(resolve, typingTime));
        await chat.clearState();
    } catch (error) {}
};

// Ambil nama admin yang terakhir mengedit spreadsheet.
// Cukup pakai scope drive.readonly yang sudah ada (tidak perlu izin tambahan).
//
// CATATAN PERBAIKAN: sebelumnya fungsi ini pakai drive.revisions.list() yang di-paginate
// sampai halaman terakhir untuk nemuin revisi paling baru. Untuk spreadsheet yang sering
// banget diedit (riwayat revisi bisa numpuk ratusan/ribuan), paginate ini kadang lambat atau
// gagal (timeout/rate limit) - itu penyebab "diubah oleh" kadang gak muncul. Sekarang pakai
// drive.files.get() yang punya field modifiedTime & lastModifyingUser langsung di metadata file
// itu sendiri - cuma 1 API call ringan, gak perlu paginate riwayat revisi sama sekali.
//
// Ditambah RETRY 2x dengan jeda singkat, buat nutup kasus gagal sesaat (network hiccup pas
// momen kirim notif). Kalau tetap gagal setelah retry, return null (pemanggil akan tetap
// nampilin baris "Diubah oleh" tapi dengan keterangan gagal, bukan disembunyikan diam-diam).
async function getAdminTerakhirEdit(percobaanKe = 1, fileId = SPREADSHEET_ID) {
    try {
        const authClient = await authorize();
        if (!authClient) return null;
        const drive = google.drive({ version: "v3", auth: authClient });

        const res = await drive.files.get({
            fileId: fileId,
            fields: "modifiedTime, lastModifyingUser(displayName, emailAddress)",
            supportsAllDrives: true,
        });

        const user = res.data.lastModifyingUser;
        if (!user) return null;
        return {
            nama: user.displayName || user.emailAddress || "Tidak diketahui",
            waktu: res.data.modifiedTime,
        };
    } catch (err) {
        console.log(`⚠️ Gagal ambil info admin terakhir edit (percobaan ${percobaanKe}/2):`, err.message);
        if (percobaanKe < 2) {
            await new Promise((res) => setTimeout(res, 2000));
            return getAdminTerakhirEdit(percobaanKe + 1, fileId);
        }
        return null;
    }
}

// Kirim semua revisi yang tertampung di antrian, lalu reset antrian & cooldown.
async function kirimNotifikasiRevisi() {
    if (antrianPerubahanPending.length === 0) return;

    const daftarRevisi = antrianPerubahanPending;
    antrianPerubahanPending = [];
    waktuNotifTerakhir = Date.now();

    const adminInfo = await getAdminTerakhirEdit();

    let headerNotif = `🚨 *ALARM REVISI ADMIN* 🚨\n\nTerdeteksi *${daftarRevisi.length} perubahan* pada jadwal.`;
    if (adminInfo) {
        const waktuFormat = new Date(adminInfo.waktu).toLocaleString("id-ID", {
            timeZone: "Asia/Makassar",
            dateStyle: "medium",
            timeStyle: "short",
        });
        headerNotif += `\n✏️ *Diubah oleh:* ${adminInfo.nama}`;
        headerNotif += `\n🕒 *Waktu edit:* ${waktuFormat} WITA`;
    } else {
        // Tetap tampilkan barisnya biar jelas ini emang gagal dideteksi (jaringan/API),
        // bukan diam-diam disembunyikan seolah fiturnya gak jalan.
        headerNotif += `\n✏️ *Diubah oleh:* (gagal dideteksi - coba lagi di revisi berikutnya)`;
    }
    headerNotif += `\n\nDetail lengkap di bawah ini:`;

    await kirimWA(ID_TUJUAN_NOTIFIKASI, headerNotif);
    for (const detailPerubahan of daftarRevisi) {
        await kirimWA(ID_TUJUAN_NOTIFIKASI, detailPerubahan);
    }
    await kirimWA(ID_TUJUAN_NOTIFIKASI, `💡 _Ketik *1* atau *2* untuk melihat detail peralatan terbaru._`);
}

const jalankanRonda = async () => {
    if (rondaJadwalBerjalan) return; // ronda sebelumnya belum selesai (download Excel lambat), jangan numpuk
    rondaJadwalBerjalan = true;
    console.log("🕵️ Meronda 2 Bulan Sekaligus...");
    try {
        const dataTerbaru = await getJadwalMultiBulan();
        if (!dataTerbaru || dataTerbaru.length === 0) return;

        if (objekDataLama) {
            // Sanity check tambahan: kalau jumlah event tiba-tiba anjlok drastis (>50%)
            // dibanding ronda sebelumnya, ini kemungkinan besar data yang kefetch rusak/parsial
            // (misal salah satu bulan gagal ke-parse tapi gak nge-throw error), BUKAN beneran
            // banyak event dihapus admin sekaligus. Skip ronda ini biar gak salah alarm massal.
            const jumlahEventLama = Object.keys(ekstrakStateEvent(objekDataLama)).length;
            const jumlahEventBaru = Object.keys(ekstrakStateEvent(dataTerbaru)).length;
            if (jumlahEventLama >= 5 && jumlahEventBaru < jumlahEventLama * 0.5) {
                console.log(`⚠️ Jumlah event anjlok drastis (${jumlahEventLama} -> ${jumlahEventBaru}), kemungkinan data korup/parsial. Ronda ini di-skip, objekDataLama TIDAK ditimpa.`);
                return;
            }

            if (JSON.stringify(dataTerbaru) !== JSON.stringify(objekDataLama)) {
                console.log("🔔 Ada perubahan jadwal!");
                const daftarRevisiBaru = cariPerubahanEvent(objekDataLama, dataTerbaru);
                if (daftarRevisiBaru.length > 0) {
                    antrianPerubahanPending.push(...daftarRevisiBaru);
                }
            }
        }
        objekDataLama = dataTerbaru;

        // Throttle: kirim sekarang kalau belum pernah kirim, atau cooldown-nya udah lewat.
        // Kalau belum lewat cooldown, perubahan tetap ketampung di antrian dan
        // baru dikirim gabungan begitu cooldown lewat (di ronda berikutnya).
        const sudahLewatCooldown = Date.now() - waktuNotifTerakhir >= JEDA_MINIMAL_NOTIF_MS;
        if (antrianPerubahanPending.length > 0 && sudahLewatCooldown) {
            await kirimNotifikasiRevisi();
        }

        // Cek konflik jadwal (crew bentrok loading s/d acara), dedupe otomatis
        await cekDanNotifKonflikJadwal(dataTerbaru);
    } catch (err) {
        console.log(`❌ Sistem Ronda Gagal (objekDataLama TIDAK ditimpa): ${err.message}`);
        sendSystemAlert(`❌ Sistem Ronda Gagal: ${err.message}`);
    } finally {
        rondaJadwalBerjalan = false;
    }
};

function ekstrakStateEvent(rawData) {
    let state = {};
    if (!rawData || !Array.isArray(rawData)) return state;
    let blocks = [];
    let currentBlock = [];
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        const colA = row && row[0] ? row[0].toString().trim() : "";
        if (/^\d+$/.test(colA)) {
            if (currentBlock.length > 0) blocks.push(currentBlock);
            currentBlock = [row];
        } else if (currentBlock.length > 0) currentBlock.push(row);
    }
    if (currentBlock.length > 0) blocks.push(currentBlock);

    for (const block of blocks) {
        // masterTanggal = nomor hari di kolom A (posisi baris/slot di sheet).
        // Dipakai sebagai KUNCI STABIL agar perubahan tanggal event tetap terdeteksi
        // sebagai "perubahan", bukan dianggap event baru.
        const masterTanggal = block[0][0] ? block[0][0].toString().trim() : "";

        for (let c = 2; c < 50; c += 8) {
            const getVal = (r, col) => block[r] && block[r][col] ? block[r][col].toString().trim() : "";
            if (getVal(1, c).toUpperCase() !== "NAME") continue;
            const picName = cleanStr(getVal(1, c + 1));
            const companyName = cleanStr(getVal(2, c + 1));
            let eventTitle = cleanStr(getVal(2, c + 6));
            const venue = cleanStr(getVal(3, c + 6));
            if (!picName && !companyName && !eventTitle && !venue) continue;

            let namaTampil = eventTitle || venue || companyName || picName || "Event Tanpa Nama";
            let dateStrRaw = getVal(1, c + 6);
            let dateStr = formatTanggalExcel(dateStrRaw);
            let loadingStrRaw = getVal(4, c + 6);
            let loadingStr = loadingStrRaw || "-";
            let eventKey = `SLOT_${masterTanggal}_${c}`;

            // Objek Date asli buat perbandingan kalender (cek konflik jadwal, cek event lewat/belum)
            let tanggalObj = parseTanggalKeDate(dateStrRaw);
            let loadingObj = parseTanggalKeDate(
                loadingStrRaw,
                tanggalObj ? tanggalObj.getMonth() + 1 : null,
                tanggalObj ? tanggalObj.getFullYear() : null
            ) || tanggalObj; // kalau loading date gak jelas, anggap sama dengan tanggal event

            let crewList = [];
            let statusEvent = "";
            let itemList = [];
            let isiEventLengkap = [];
            for (let i = 1; i < block.length; i++) {
                let barisString = "";
                for (let k = c; k <= c + 7; k++) barisString += getVal(i, k) + "|";
                isiEventLengkap.push(barisString);

                if (i >= 8) {
                    let teksBaris = barisString.toUpperCase();
                    if (teksBaris.includes("CUSTOMER DETILS|")) break;
                    if (teksBaris.includes("STATUS|")) {
                        let stat = getVal(i, c + 1);
                        if (!stat || stat === "-" || stat.toUpperCase() === "STATUS") stat = getVal(i, c + 6);
                        if (stat && stat !== "-") statusEvent = stat;
                        break;
                    }
                    const crew = getVal(i, c + 6);
                    if (crew && crew !== "-" && crew.toUpperCase() !== "CREW" && crew !== "") {
                        if (crew.toUpperCase() === "DONE" || crew.toUpperCase() === "CANCEL" || crew.toUpperCase() === "CANCELLED") {
                            statusEvent = crew;
                        } else if (namaCrewValid(crew) && !crewList.includes(crew)) {
                            crewList.push(crew);
                        }
                    }

                    // Ambil data barang/alat di baris yang sama, supaya bisa dibandingkan (ditambah/dihapus)
                    const item = getVal(i, c + 1);
                    const spec = getVal(i, c + 2);
                    const qty = getVal(i, c + 3);
                    const freq = getVal(i, c + 5);
                    if (qty && item && item.toUpperCase() !== "ITEM") {
                        let namaLengkap = `${item} ${spec}`.trim().replace(/\s+/g, " ");
                        let teksAlat = `${qty} ${namaLengkap}`.trim();
                        if (freq && freq !== "-") teksAlat += ` (${freq})`;
                        itemList.push(teksAlat);
                    }
                }
            }
            state[eventKey] = {
                nama: namaTampil,
                tanggal: dateStr,
                tanggalObj: tanggalObj,
                loading: loadingStr,
                loadingObj: loadingObj,
                venue: venue,
                company: companyName,
                crew: crewList,
                status: statusEvent,
                items: itemList,
                hash: isiEventLengkap.join("~"),
            };
        }
    }
    return state;
}

// Bandingkan 2 array string, hasilkan { ditambah: [...], dihapus: [...] }
function diffArray(arrLama, arrBaru) {
    const ditambah = arrBaru.filter((x) => !arrLama.includes(x));
    const dihapus = arrLama.filter((x) => !arrBaru.includes(x));
    return { ditambah, dihapus };
}

function cariPerubahanEvent(dataLama, dataBaru) {
    let stateLama = ekstrakStateEvent(dataLama);
    let stateBaru = ekstrakStateEvent(dataBaru);
    let hasilPerubahan = [];

    // 1. Cek slot yang ada di data terbaru (event baru / event yang berubah)
    for (let key in stateBaru) {
        let baru = stateBaru[key];
        let lama = stateLama[key];

        // Slot ini sebelumnya kosong / belum tercatat -> anggap Event Baru
        if (!lama) {
            let msg = `🆕 *EVENT BARU*\n📌 *${baru.nama}*\n📅 Tanggal: ${baru.tanggal}`;
            if (baru.venue && baru.venue !== "-") msg += `\n📍 Venue: ${baru.venue}`;
            msg += baru.crew.length > 0 ? `\n👥 Crew: ${baru.crew.join(", ")}` : `\n👥 Crew: (Belum diplot)`;
            if (baru.items.length > 0) msg += `\n📦 Barang: ${baru.items.length} item terpasang`;
            hasilPerubahan.push(msg);
            continue;
        }

        if (lama.hash === baru.hash) continue; // benar-benar tidak ada perubahan

        let detail = [];

        if (lama.nama !== baru.nama) {
            detail.push(`📝 *Nama/Judul Event:*\n   "${lama.nama}" ➡️ "${baru.nama}"`);
        }
        if (lama.tanggal !== baru.tanggal) {
            detail.push(`📅 *Tanggal Event:*\n   ${lama.tanggal} ➡️ ${baru.tanggal}`);
        }
        if (lama.loading !== baru.loading) {
            detail.push(`🚚 *Tanggal Loading:*\n   ${lama.loading} ➡️ ${baru.loading}`);
        }
        if (lama.venue !== baru.venue) {
            detail.push(`📍 *Venue:*\n   ${lama.venue || "-"} ➡️ ${baru.venue || "-"}`);
        }
        if (lama.company !== baru.company) {
            detail.push(`🏢 *Client/Company:*\n   ${lama.company || "-"} ➡️ ${baru.company || "-"}`);
        }
        if (lama.status !== baru.status) {
            detail.push(`🏷️ *Status:*\n   ${lama.status || "(kosong)"} ➡️ ${baru.status || "(kosong)"}`);
        }

        const crewDiff = diffArray(lama.crew, baru.crew);
        if (crewDiff.ditambah.length > 0 || crewDiff.dihapus.length > 0) {
            let crewMsg = `👥 *Crew berubah:*`;
            if (crewDiff.dihapus.length > 0) crewMsg += `\n   ➖ Dicopot: ${crewDiff.dihapus.join(", ")}`;
            if (crewDiff.ditambah.length > 0) crewMsg += `\n   ➕ Ditugaskan: ${crewDiff.ditambah.join(", ")}`;
            detail.push(crewMsg);
        }

        const itemDiff = diffArray(lama.items, baru.items);
        if (itemDiff.ditambah.length > 0 || itemDiff.dihapus.length > 0) {
            let itemMsg = `📦 *Barang berubah:*`;
            if (itemDiff.dihapus.length > 0) itemMsg += `\n   ➖ Dihapus: ${itemDiff.dihapus.join(", ")}`;
            if (itemDiff.ditambah.length > 0) itemMsg += `\n   ➕ Ditambahkan: ${itemDiff.ditambah.join(", ")}`;
            detail.push(itemMsg);
        }

        if (detail.length === 0) {
            detail.push(`ℹ️ Ada perubahan kecil pada data mentah (misal spasi/format), tidak signifikan.`);
        }

        let msg = `📌 *${baru.nama}* (${baru.tanggal})\n` + detail.join("\n");
        hasilPerubahan.push(msg);
    }

    // 2. Cek slot yang hilang di data terbaru (event dihapus dari sheet)
    for (let key in stateLama) {
        if (!stateBaru[key]) {
            let lama = stateLama[key];
            hasilPerubahan.push(`🗑️ *EVENT DIHAPUS*\n📌 *${lama.nama}* (${lama.tanggal}) sudah tidak ada lagi di sheet.`);
        }
    }

    return hasilPerubahan;
}

// ============================================================================
// 7. DETEKSI KONFLIK JADWAL (DOUBLE BOOKING) & PENCARIAN CREW/ALAT
// ============================================================================

// Cek apakah ada CREW yang ke-assign ke 2 event yang jadwalnya beririsan (loading s/d acara).
// Contoh: Event A loading tgl 1, acara tgl 2. Event B loading & acara tgl 2.
// -> beririsan di tgl 2, kalau crew-nya sama maka ini baru dianggap bentrok.
// Alat/barang sengaja TIDAK dicek lagi di sini (kebanyakan noise), fokus ke crew saja.
// Event yang tanggal acaranya sudah lewat hari ini otomatis di-skip (gak relevan lagi).
function cekKonflikJadwal(rawData) {
    const state = ekstrakStateEvent(rawData);
    const hariIni = new Date();
    hariIni.setHours(0, 0, 0, 0);

    // Ambil event yang tanggalnya bisa ditentukan DAN belum lewat
    const daftarEvent = Object.values(state).filter((ev) => {
        if (!ev.tanggalObj) return false; // tanggal gak bisa dibaca -> skip biar gak salah alarm
        return ev.tanggalObj.getTime() >= hariIni.getTime();
    });

    let daftarKonflik = [];
    for (let i = 0; i < daftarEvent.length; i++) {
        for (let j = i + 1; j < daftarEvent.length; j++) {
            const a = daftarEvent[i];
            const b = daftarEvent[j];

            const mulaiA = a.loadingObj || a.tanggalObj;
            const selesaiA = a.tanggalObj;
            const mulaiB = b.loadingObj || b.tanggalObj;
            const selesaiB = b.tanggalObj;

            // Overlap range [mulaiA, selesaiA] vs [mulaiB, selesaiB]
            const beririsan = mulaiA.getTime() <= selesaiB.getTime() && mulaiB.getTime() <= selesaiA.getTime();
            if (!beririsan) continue;

            const crewBentrok = a.crew.filter((c) => b.crew.some((c2) => c2.toUpperCase() === c.toUpperCase()));
            for (const crew of crewBentrok) {
                daftarKonflik.push(
                    `👤 *${crew}* ke-assign bareng di 2 event yang jadwalnya beririsan:\n` +
                    `   • ${a.nama} (Loading ${a.loading} ➡️ Acara ${a.tanggal})\n` +
                    `   • ${b.nama} (Loading ${b.loading} ➡️ Acara ${b.tanggal})`
                );
            }
        }
    }
    return daftarKonflik;
}

// Bandingkan hasil cek konflik dengan hasil sebelumnya, kirim notif hanya untuk yang BARU
// (biar gak spam notif yang sama tiap ronda selama konfliknya belum diselesaikan).
async function cekDanNotifKonflikJadwal(rawData) {
    const konflikSekarang = cekKonflikJadwal(rawData);
    const konflikBaru = konflikSekarang.filter((k) => !daftarKonflikSebelumnya.includes(k));
    daftarKonflikSebelumnya = konflikSekarang;

    if (konflikBaru.length > 0) {
        await kirimWA(ID_TUJUAN_NOTIFIKASI, `⚠️ *KONFLIK JADWAL TERDETEKSI* ⚠️\n\nDitemukan *${konflikBaru.length} bentrok baru*:`);
        for (const k of konflikBaru) {
            await kirimWA(ID_TUJUAN_NOTIFIKASI, k);
        }
    }
}

// Cari semua event yang melibatkan crew tertentu (pencocokan sebagian/contains, case-insensitive)
function cariJadwalByCrew(rawData, namaCrew) {
    const state = ekstrakStateEvent(rawData);
    const target = namaCrew.toUpperCase();
    let hasil = [];
    for (const key in state) {
        const ev = state[key];
        if (ev.crew.some((c) => c.toUpperCase().includes(target))) hasil.push(ev);
    }
    return hasil;
}

// Cari semua event yang memakai alat/barang tertentu (pencocokan sebagian/contains, case-insensitive)
function cariJadwalByAlat(rawData, namaAlat) {
    const state = ekstrakStateEvent(rawData);
    const target = namaAlat.toUpperCase();
    let hasil = [];
    for (const key in state) {
        const ev = state[key];
        if (ev.items.some((i) => i.toUpperCase().includes(target))) hasil.push(ev);
    }
    return hasil;
}

client.on("qr", (qr) => qrcode.generate(qr, { small: true }));

client.on("ready", async () => {
    console.log("✅ Bot Siap!");
    isBotReady = true;
    jumlahRetryReconnect = 0; // reset counter tiap kali berhasil connect

    try {
        objekDataLama = await getJadwalMultiBulan();
    } catch (err) {
        console.log("⚠️ Gagal ambil jadwal awal saat startup, akan dicoba lagi di ronda berikutnya:", err.message);
    }

    // Event "ready" bisa kepanggil LAGI setelah reconnect otomatis. Tanpa guard ini, setiap reconnect
    // menambah setInterval & cron baru -> notifikasi jadi dobel/triple. Daftarkan penjadwal cuma SEKALI.
    if (!penjadwalSudahAktif) {
        penjadwalSudahAktif = true;

        setInterval(jalankanRonda, WAKTU_RONDA_MS);

        // Ronda payroll: muat baseline dari disk, jalankan sekali (baseline/diff), lalu tiap 5 menit
        await muatSnapshotGaji();
        jalankanRondaGaji();
        setInterval(jalankanRondaGaji, WAKTU_RONDA_GAJI_MS);

        cron.schedule('0 6 * * *', async () => {
            try {
                const dateObj = new Date();
                const tglHariIni = dateObj.getDate().toString();
                const freshData = await getJadwalMultiBulan();
                objekDataLama = freshData;

                const daftarPesan = prosesDataKePesanWA(objekDataLama, tglHariIni, "");
                let sapaanPagi = `🌅 *MORNING BRIEFING*\nSelamat pagi! Hari ini ada *${daftarPesan.length} Event* yang tercatat di sistem.`;
                await kirimWA(ID_TUJUAN_NOTIFIKASI, sapaanPagi);

                for (const pesan of daftarPesan) {
                    await kirimWA(ID_TUJUAN_NOTIFIKASI, pesan);
                }
            } catch (error) {
                sendSystemAlert(`❌ Gagal mengirim Morning Briefing: ${error.message}`);
            }
        }, { scheduled: true, timezone: "Asia/Makassar" });

        // Pengingat pencairan gaji, jam 09:00 WITA (hanya terkirim di tanggal pengingat & akhir bulan)
        cron.schedule('0 9 * * *', async () => {
            try {
                await kirimPengingatGaji();
            } catch (error) {
                console.log("⚠️ Gagal kirim pengingat gaji:", error.message);
            }
        }, { scheduled: true, timezone: "Asia/Makassar" });
    }

    sendSystemAlert("✅ AGS Bot Enterprise System Online!");
});

client.on("disconnected", async (reason) => {
    console.log("❌ Bot terputus!", reason);
    isBotReady = false;

    // WA lagi down, jadi notif ini WAJIB lewat Telegram, bukan lewat WA itu sendiri
    await kirimTelegramFallback(`⚠️ *WA Bot Terputus*\nAlasan: ${reason}\n\nMencoba reconnect otomatis...`);

    // Kalau sesi di-logout manual/dihapus, retry gak ada gunanya - butuh scan QR ulang di server
    if (reason === "LOGOUT") {
        await kirimTelegramFallback(`🚪 Sesi WA logout. Perlu scan QR ulang manual di server, proses akan berhenti.`);
        process.exit(1);
        return;
    }

    if (jumlahRetryReconnect < MAX_RETRY_RECONNECT) {
        jumlahRetryReconnect++;
        const jedaRetry = 5000 * jumlahRetryReconnect; // backoff makin lama tiap percobaan
        console.log(`🔄 Percobaan reconnect ke-${jumlahRetryReconnect}/${MAX_RETRY_RECONNECT} dalam ${jedaRetry / 1000}s...`);
        setTimeout(() => {
            client.initialize().catch((err) => {
                console.log("❌ Reconnect gagal:", err.message);
            });
        }, jedaRetry);
    } else {
        await kirimTelegramFallback(`🛑 Reconnect otomatis gagal ${MAX_RETRY_RECONNECT}x. Proses akan di-restart lewat PM2.`);
        process.exit(1);
    }
});

process.on("unhandledRejection", (error) => {
    sendSystemAlert(`⚠️ *CRITICAL ERROR* Node.js:\n${error.message}`);
});

client.on("message", async (msg) => {
    const ALLOWED_ADMINS = ["628970282769", "35270472773718"];
    const contact = await msg.getContact();
    const nomorPengirim = contact.number; 

    if (!nomorPengirim || !ALLOWED_ADMINS.includes(nomorPengirim)) return;

    const text = msg.body.toLowerCase().trim();
    
    let chat = null;
    try { chat = await msg.getChat(); } catch (error) {}

    const balasPesan = async (teksBalasan) => {
        if (chat) await simulateTyping(chat, teksBalasan);
        try { await msg.reply(teksBalasan); } 
        catch (err) { await client.sendMessage(ID_TUJUAN_NOTIFIKASI, teksBalasan); }
    };

    if (["halo", "menu", "jadwal", "bot"].includes(text)) {
        const balasanMenu = `━━━━━━━━━━━━━━━━━━\n🤖 *AGS ENTERPRISE BOT*\n━━━━━━━━━━━━━━━━━━\n\n1️⃣ 📍 Jadwal Hari Ini\n2️⃣ 📍 Jadwal Besok\n3️⃣ 📆 Semua Jadwal Mendatang\n\n💰 *Gaji Pribadi:*\n• \`gaji\` (Bulan Berjalan)\n• \`gaji [bulan]\` (Contoh: \`gaji mei\`)\n• \`rekap tahun\` (Ringkasan Tahunan)\n• \`top venue\` (Klien Terbesar)\n• \`proyeksi tabungan\` (Estimasi Saldo)\n\n🏢 *Payroll Seluruh Karyawan:*\n• \`gaji karyawan [nama]\` (Slip siapa aja, + bulan opsional)\n• \`daftar karyawan\` (Daftar nama & tim)\n• \`status gaji\` (Periode mana yang sudah cair)\n• \`rekap payroll\` (Total cash-out + per periode)\n• \`kasbon aktif\` (Daftar kasbon)\n• \`leaderboard gaji\` (Top 10 fee & job)\n• \`bandingkan gaji [nama]\` (Bulan ini vs lalu)\n• \`tabungan [nama]\` / \`tabungan semua\`\n\n🔍 *Pencarian Cerdas:*\nKetik \`cari [kata kunci]\`\n\n👥 *Cek Jadwal:*\n• \`crew [nama]\` (Semua job crew tsb)\n• \`alat [nama]\` (Semua event yang pakai alat tsb)\n• \`cek konflik\` (Crew bentrok jadwal)\n\n🔔 _Bot otomatis mengabari kalau jadwal/payroll berubah & saat gaji cair._\n\n✏️ Ketik pilihan Anda...`;
        await balasPesan(balasanMenu);
    } 
    
    else if (text === "gaji" || text === "slip gaji") {
        await balasPesan(`⏳ Mengakses brankas data slip gaji bulan ini...`);
        const hasilGaji = await parserEngineSlipGaji(NAMA_SAYA_DI_ABSEN, null); 
        await balasPesan(hasilGaji);
    }

    // ---- PAYROLL SELURUH KARYAWAN ----
    // PENTING: blok ini HARUS di atas branch `gaji [bulan]` (pribadi). Kalau di bawahnya,
    // "gaji karyawan budi" ketangkap duluan sebagai arsip bulan "karyawan budi".
    else if (text.startsWith("gaji karyawan")) {
        const sisaTeks = text.replace("gaji karyawan", "").trim();
        if (!sisaTeks) return balasPesan("⚠️ Format: `gaji karyawan [nama]` atau `gaji karyawan [nama] [bulan]`.\nKetik `daftar karyawan` untuk lihat nama yang terdaftar.");

        const kata = sisaTeks.split(/\s+/);
        let bulanQuery = null;
        if (kata.length > 1 && NAMA_BULAN_LOWER.includes(kata[kata.length - 1])) bulanQuery = kata.pop();
        const namaKaryawan = kata.join(" ").trim();
        if (namaKaryawan.length < 2) return balasPesan("⚠️ Nama karyawan terlalu pendek (minimal 2 huruf).");

        await balasPesan(`⏳ Mengakses slip gaji *${namaKaryawan.toUpperCase()}*${bulanQuery ? ` (${bulanQuery.toUpperCase()})` : ""}...`);
        await balasPesan(await getSlipGajiKaryawan(namaKaryawan, bulanQuery));
    }

    else if (text === "daftar karyawan" || text.startsWith("daftar karyawan ")) {
        const bulanQuery = text.replace("daftar karyawan", "").trim() || null;
        await balasPesan(await getDaftarKaryawan(bulanQuery));
    }

    else if (text.startsWith("bandingkan gaji")) {
        const nama = text.replace("bandingkan gaji", "").trim() || NAMA_SAYA_DI_ABSEN;
        await balasPesan(`⏳ Membandingkan gaji *${nama.toUpperCase()}* bulan ini vs bulan lalu...`);
        await balasPesan(await getPerbandinganGaji(nama));
    }

    else if (text === "tabungan" || text.startsWith("tabungan ")) {
        const sisa = text.replace("tabungan", "").trim();
        await balasPesan(`⏳ Membuka data tabungan...`);
        if (sisa === "semua") await balasPesan(await getTabunganKaryawan(null));
        else await balasPesan(await getTabunganKaryawan(sisa || NAMA_SAYA_DI_ABSEN));
    }

    else if (text === "status gaji" || text === "cek gaji masuk") {
        await balasPesan(`⏳ Mengecek status pencairan gaji...`);
        await balasPesan(await getStatusGaji());
    }

    else if (text === "rekap payroll" || text.startsWith("rekap payroll ")) {
        const bulanQuery = text.replace("rekap payroll", "").trim() || null;
        await balasPesan(`⏳ Menghitung rekap payroll seluruh karyawan${bulanQuery ? ` (${bulanQuery.toUpperCase()})` : " bulan ini"}...`);
        await balasPesan(await getRekapPayrollBulanIni(bulanQuery));
    }

    else if (text === "kasbon aktif" || text === "cek kasbon" || text.startsWith("kasbon aktif ")) {
        const bulanQuery = text.replace("kasbon aktif", "").replace("cek kasbon", "").trim() || null;
        await balasPesan(`⏳ Mengecek daftar kasbon aktif${bulanQuery ? ` (${bulanQuery.toUpperCase()})` : " bulan ini"}...`);
        await balasPesan(await getDaftarKasbonAktif(bulanQuery));
    }

    else if (text === "leaderboard gaji" || text === "leaderboard payroll" || text.startsWith("leaderboard gaji ") || text.startsWith("leaderboard payroll ")) {
        const bulanQuery = text.replace("leaderboard gaji", "").replace("leaderboard payroll", "").trim() || null;
        await balasPesan(`⏳ Menyusun leaderboard top 10${bulanQuery ? ` (${bulanQuery.toUpperCase()})` : " bulan ini"}...`);
        await balasPesan(await getLeaderboardPayroll(bulanQuery));
    }

    // ---- GAJI PRIBADI (spreadsheet lama) ----
    else if (text.startsWith("gaji ")) {
        let bulanQuery = text.replace("gaji ", "").trim();
        await balasPesan(`⏳ Membuka arsip slip gaji bulan *${bulanQuery.toUpperCase()}*...`);
        const hasilArsip = await parserEngineSlipGaji(NAMA_SAYA_DI_ABSEN, bulanQuery);
        await balasPesan(hasilArsip);
    }

    else if (text === "rekap tahun" || text === "rekap tahunan") {
        await balasPesan(`⏳ Menghitung rekapitulasi finansial tahunan...`);
        const hasilRekap = await getRekapTahunan(NAMA_SAYA_DI_ABSEN);
        await balasPesan(hasilRekap);
    }

    else if (text === "top venue" || text === "venue termahal") {
        await balasPesan(`⏳ Menganalisis venue dan klien penghasilan terbesar...`);
        const hasilTop = await getTopVenue(NAMA_SAYA_DI_ABSEN);
        await balasPesan(hasilTop);
    }

    else if (text === "proyeksi tabungan") {
        await balasPesan(`⏳ Menghitung proyeksi saldo tabungan akhir tahun...`);
        const hasilProyeksi = await getProyeksiTabungan(NAMA_SAYA_DI_ABSEN);
        await balasPesan(hasilProyeksi);
    }

    else if (text === "cek konflik" || text === "konflik jadwal") {
        await balasPesan(`⏳ Mengecek bentrokan jadwal crew (loading s/d acara)...`);
        const dataCek = await ambilDataJadwal();
        const hasilKonflik = cekKonflikJadwal(dataCek);
        if (hasilKonflik.length === 0) {
            await balasPesan(`✅ Tidak ada bentrokan jadwal crew yang terdeteksi (untuk event hari ini & ke depan).`);
        } else {
            await balasPesan(`⚠️ Ditemukan *${hasilKonflik.length} bentrokan*:`);
            for (const k of hasilKonflik) {
                await balasPesan(k);
                await new Promise((res) => setTimeout(res, 500));
            }
        }
    }

    else if (text.startsWith("crew ")) {
        const namaCrew = text.replace("crew ", "").trim();
        if (namaCrew.length < 2) return balasPesan("⚠️ Nama crew terlalu pendek. Minimal 2 huruf.");

        await balasPesan(`⏳ Mencari jadwal crew *${namaCrew.toUpperCase()}*...`);
        const dataCek = await ambilDataJadwal();
        const hasil = cariJadwalByCrew(dataCek, namaCrew);

        if (hasil.length === 0) {
            await balasPesan(`ℹ️ Tidak ada jadwal ditemukan untuk crew *${namaCrew.toUpperCase()}*.`);
        } else {
            await balasPesan(`✅ Ditemukan *${hasil.length} job* untuk crew *${namaCrew.toUpperCase()}*:`);
            for (const ev of hasil) {
                let msg = `📌 *${ev.nama}*\n📅 ${ev.tanggal}\n📍 ${ev.venue || "-"}`;
                if (ev.status) msg += `\n🏷️ Status: ${ev.status}`;
                await balasPesan(msg);
                await new Promise((res) => setTimeout(res, 500));
            }
        }
    }

    else if (text.startsWith("alat ")) {
        const namaAlat = text.replace("alat ", "").trim();
        if (namaAlat.length < 2) return balasPesan("⚠️ Nama alat terlalu pendek. Minimal 2 huruf.");

        await balasPesan(`⏳ Mencari jadwal pemakaian alat *${namaAlat.toUpperCase()}*...`);
        const dataCek = await ambilDataJadwal();
        const hasil = cariJadwalByAlat(dataCek, namaAlat);

        if (hasil.length === 0) {
            await balasPesan(`ℹ️ Tidak ada jadwal ditemukan untuk alat *${namaAlat.toUpperCase()}*.`);
        } else {
            await balasPesan(`✅ Alat *${namaAlat.toUpperCase()}* dipakai di *${hasil.length} event*:`);
            for (const ev of hasil) {
                const itemCocok = ev.items.filter((i) => i.toUpperCase().includes(namaAlat.toUpperCase()));
                let msg = `📌 *${ev.nama}*\n📅 ${ev.tanggal}\n📦 ${itemCocok.join(", ")}`;
                await balasPesan(msg);
                await new Promise((res) => setTimeout(res, 500));
            }
        }
    }

    else if (text.startsWith("cari ") || text.startsWith("search ")) {
        let keyword = text.replace("cari ", "").replace("search ", "").trim();
        if (keyword.length < 3) return balasPesan("⚠️ Kata kunci terlalu pendek. Minimal 3 huruf.");

        let targetTanggal = "";
        let infoWaktu = "Mendatang (Disaring Otomatis)";

        if (keyword.endsWith(" hari ini")) {
            keyword = keyword.replace(" hari ini", "").trim();
            targetTanggal = new Date().getDate().toString();
            infoWaktu = "Hari Ini";
        } else if (keyword.endsWith(" besok")) {
            keyword = keyword.replace(" besok", "").trim();
            let besok = new Date(); 
            besok.setDate(besok.getDate() + 1);
            targetTanggal = besok.getDate().toString();
            infoWaktu = "Besok";
        }

        await balasPesan(`⏳ Mencari *"${keyword}"* untuk jadwal ${infoWaktu}...`);

        const dataCache = await ambilDataJadwal();
        const daftarPesan = prosesDataKePesanWA(dataCache, targetTanggal, keyword);

        if (daftarPesan.length === 0) {
            await balasPesan(`ℹ️ Tidak ditemukan hasil untuk *"${keyword}"* (${infoWaktu}).`);
        } else {
            await balasPesan(`✅ Ditemukan *${daftarPesan.length} Hasil* pencarian:`);
            for (const pesan of daftarPesan) {
                await balasPesan(pesan);
                await new Promise((res) => setTimeout(res, 500)); 
            }
        }
    }

    else if (["1", "2", "3"].includes(text)) {
        const dateObj = new Date();
        let tglTarget = "";
        let labelTarget = "";

        if (text === "1") {
            tglTarget = dateObj.getDate().toString();
            labelTarget = "Hari Ini";
        } else if (text === "2") {
            dateObj.setDate(dateObj.getDate() + 1);
            tglTarget = dateObj.getDate().toString();
            labelTarget = "Besok";
        } else if (text === "3") {
            tglTarget = "";
            labelTarget = "Semua Jadwal Mendatang";
        }

        await balasPesan(`⚡ Menarik data ${labelTarget} (Cache Mode)...`);

        const dataUntukPerintah = await ambilDataJadwal();
        const keywordVirtual = text === "3" ? "[SHOW_ALL]" : ""; 
        const daftarPesan = prosesDataKePesanWA(dataUntukPerintah, tglTarget, keywordVirtual);

        if (daftarPesan.length === 0) {
            await balasPesan(`ℹ️ Tidak ada jadwal untuk ${labelTarget}.`);
        } else {
            for (const pesan of daftarPesan) {
                const pesanBersih = pesan.replace("[SHOW_ALL]", "").trim(); 
                await balasPesan(pesanBersih);
                await new Promise((res) => setTimeout(res, 500)); 
            }
        }
    }
});

client.initialize();