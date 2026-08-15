const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const { google } = require("googleapis");
const fs = require("fs").promises;
const readline = require("readline");
const XLSX = require("xlsx");
const cron = require("node-cron");

// ============================================================================
// 1. KONFIGURASI UTAMA
// ============================================================================
const SPREADSHEET_ID = "1DLcMkga8UiRtRJ3ZQIPMRQb-5d1IFiu_"; // Excel Jadwal
const SPREADSHEET_ID_GAJI = "19QJEZ11K63KYfGfiT4ECSHodfsa_KgCpPz7gDg3UWxk"; // Excel Gaji
const NAMA_SAYA_DI_ABSEN = "JAYAK"; // Sesuai tulisan di absen admin

const ID_TUJUAN_NOTIFIKASI = "628970282769@c.us";
const WAKTU_RONDA_MS = 1 * 60 * 1000; // 1 Menit

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
async function getJadwalDariExcel(targetDateObj) {
    const authClient = await authorize();
    if (!authClient) return null;
    const drive = google.drive({ version: "v3", auth: authClient });

    try {
        const res = await drive.files.get({ fileId: SPREADSHEET_ID, alt: "media" }, { responseType: "arraybuffer" });
        const workbook = XLSX.read(res.data, { type: "buffer" });
        const namaBulan = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
        const targetSheetName = `${namaBulan[targetDateObj.getMonth()]} ${targetDateObj.getFullYear()}`;
        
        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet) return null;

        return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    } catch (error) {
        return null;
    }
}

async function getJadwalMultiBulan() {
    const dateIni = new Date();
    const dateDepan = new Date();
    dateDepan.setMonth(dateDepan.getMonth() + 1);
    
    const dataIni = await getJadwalDariExcel(dateIni) || [];
    const dataDepan = await getJadwalDariExcel(dateDepan) || [];
    return [...dataIni, ...dataDepan];
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
// 6. ENGINE WHATSAPP & BOT LOGIC
// ============================================================================
const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"] },
});

const sendSystemAlert = async (text) => {
    console.log(text);
    if (isBotReady) {
        try { await client.sendMessage(ID_TUJUAN_NOTIFIKASI, text); } catch (e) {}
    }
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

// Ambil nama admin yang terakhir mengedit spreadsheet, via Google Drive Revisions API.
// Cukup pakai scope drive.readonly yang sudah ada (tidak perlu izin tambahan).
// Catatan: ini atribusi per-FILE (siapa yang terakhir nyentuh file), bukan per-cell.
async function getAdminTerakhirEdit() {
    try {
        const authClient = await authorize();
        if (!authClient) return null;
        const drive = google.drive({ version: "v3", auth: authClient });

        let revisiTerakhir = null;
        let pageToken = null;
        do {
            const res = await drive.revisions.list({
                fileId: SPREADSHEET_ID,
                fields: "nextPageToken, revisions(id, modifiedTime, lastModifyingUser(displayName, emailAddress))",
                pageSize: 1000,
                pageToken: pageToken || undefined,
            });
            const revisions = res.data.revisions || [];
            if (revisions.length > 0) revisiTerakhir = revisions[revisions.length - 1];
            pageToken = res.data.nextPageToken;
        } while (pageToken);

        if (!revisiTerakhir || !revisiTerakhir.lastModifyingUser) return null;
        return {
            nama: revisiTerakhir.lastModifyingUser.displayName || revisiTerakhir.lastModifyingUser.emailAddress || "Tidak diketahui",
            waktu: revisiTerakhir.modifiedTime,
        };
    } catch (err) {
        console.log("⚠️ Gagal ambil riwayat revisi (info admin dilewati):", err.message);
        return null;
    }
}

const jalankanRonda = async () => {
    console.log("🕵️ Meronda 2 Bulan Sekaligus...");
    try {
        const dataTerbaru = await getJadwalMultiBulan();
        if (!dataTerbaru || dataTerbaru.length === 0) return;

        if (objekDataLama && JSON.stringify(dataTerbaru) !== JSON.stringify(objekDataLama)) {
            console.log("🔔 Ada perubahan jadwal!");
            const daftarRevisi = cariPerubahanEvent(objekDataLama, dataTerbaru);
            if (daftarRevisi.length > 0) {
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
                }
                headerNotif += `\n\nDetail lengkap di bawah ini:`;

                await client.sendMessage(ID_TUJUAN_NOTIFIKASI, headerNotif);

                for (const detailPerubahan of daftarRevisi) {
                    await new Promise((res) => setTimeout(res, 800));
                    await client.sendMessage(ID_TUJUAN_NOTIFIKASI, detailPerubahan);
                }

                await new Promise((res) => setTimeout(res, 500));
                await client.sendMessage(ID_TUJUAN_NOTIFIKASI, `💡 _Ketik *1* atau *2* untuk melihat detail peralatan terbaru._`);
            }
        }
        objekDataLama = dataTerbaru;
    } catch (err) {
        sendSystemAlert(`❌ Sistem Ronda Gagal: ${err.message}`);
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
            let dateStr = formatTanggalExcel(getVal(1, c + 6));
            let loadingStr = getVal(4, c + 6) || "-";
            let eventKey = `SLOT_${masterTanggal}_${c}`;

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
                        } else if (!crewList.includes(crew)) crewList.push(crew);
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
                loading: loadingStr,
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

client.on("qr", (qr) => qrcode.generate(qr, { small: true }));

client.on("ready", async () => {
    console.log("✅ Bot Siap!");
    isBotReady = true;

    objekDataLama = await getJadwalMultiBulan();
    setInterval(jalankanRonda, WAKTU_RONDA_MS);

    cron.schedule('0 6 * * *', async () => {
        try {
            const dateObj = new Date();
            const tglHariIni = dateObj.getDate().toString();
            const freshData = await getJadwalMultiBulan(); 
            if (freshData) objekDataLama = freshData;

            const daftarPesan = prosesDataKePesanWA(objekDataLama, tglHariIni, "");
            let sapaanPagi = `🌅 *MORNING BRIEFING*\nSelamat pagi! Hari ini ada *${daftarPesan.length} Event* yang tercatat di sistem.`;
            await client.sendMessage(ID_TUJUAN_NOTIFIKASI, sapaanPagi);

            for (const pesan of daftarPesan) {
                await new Promise((res) => setTimeout(res, 1000));
                await client.sendMessage(ID_TUJUAN_NOTIFIKASI, pesan);
            }
        } catch (error) {
            sendSystemAlert(`❌ Gagal mengirim Morning Briefing: ${error.message}`);
        }
    }, { scheduled: true, timezone: "Asia/Makassar" });

    sendSystemAlert("✅ AGS Bot Enterprise System Online!");
});

client.on("disconnected", (reason) => {
    console.log("❌ Bot terputus!", reason);
    process.exit(1);
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
        const balasanMenu = `━━━━━━━━━━━━━━━━━━\n🤖 *AGS ENTERPRISE BOT*\n━━━━━━━━━━━━━━━━━━\n\n1️⃣ 📍 Jadwal Hari Ini\n2️⃣ 📍 Jadwal Besok\n3️⃣ 📆 Semua Jadwal Mendatang\n\n💰 *Keuangan & Slip Gaji:*\n• \`gaji\` (Bulan Berjalan)\n• \`gaji [bulan]\` (Contoh: \`gaji mei\`)\n• \`rekap tahun\` (Ringkasan Tahunan)\n• \`top venue\` (Peringkat Klien Terbesar)\n• \`proyeksi tabungan\` (Estimasi Saldo)\n\n🔍 *Pencarian Cerdas:*\nKetik \`cari [kata kunci]\`\n\n✏️ Ketik pilihan Anda...`;
        await balasPesan(balasanMenu);
    } 
    
    else if (text === "gaji" || text === "slip gaji") {
        await balasPesan(`⏳ Mengakses brankas data slip gaji bulan ini...`);
        const hasilGaji = await parserEngineSlipGaji(NAMA_SAYA_DI_ABSEN, null); 
        await balasPesan(hasilGaji);
    }

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

        const dataCache = objekDataLama || (await getJadwalMultiBulan());
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

        if (!objekDataLama) objekDataLama = await getJadwalMultiBulan();
        const keywordVirtual = text === "3" ? "[SHOW_ALL]" : ""; 
        const daftarPesan = prosesDataKePesanWA(objekDataLama, tglTarget, keywordVirtual);

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