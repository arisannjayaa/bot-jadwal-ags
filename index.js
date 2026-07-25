const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const { google } = require("googleapis");
const fs = require("fs").promises;
const readline = require("readline");
const XLSX = require("xlsx");
const cron = require("node-cron"); // TAMBAHAN: Untuk fitur Daily Morning Briefing

// ============================================================================
// 1. KONFIGURASI UTAMA
// ============================================================================
const SPREADSHEET_ID = "1DLcMkga8UiRtRJ3ZQIPMRQb-5d1IFiu_";
const ID_TUJUAN_NOTIFIKASI = "628970282769@c.us";
const WAKTU_RONDA_MS = 1 * 60 * 1000; // 1 Menit

// VARIABEL CACHE (In-Memory Caching)
let objekDataLama = null; // Menyimpan raw data Excel secara global
let globalAuthClient = null;
let isBotReady = false;


// ============================================================================
// 2. SISTEM LOGIN OAUTH 2.0 (STRICT - Tidak Diubah)
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
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: "offline",
        scope: ["https://www.googleapis.com/auth/drive.readonly"],
    });
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


// ============================================================================
// 3. FUNGSI BANTUAN & FORMATTING (STRICT)
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
// 4. ENGINE PARSING GOOGLE SHEETS
// ============================================================================
async function getJadwalDariExcel(targetDateObj = new Date()) {
    const authClient = await authorize();
    if (!authClient) return null;
    const drive = google.drive({ version: "v3", auth: authClient });

    try {
        const res = await drive.files.get({ fileId: SPREADSHEET_ID, alt: "media" }, { responseType: "arraybuffer" });
        const workbook = XLSX.read(res.data, { type: "buffer" });
        const namaBulan = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
        const targetSheetName = `${namaBulan[targetDateObj.getMonth()]} ${targetDateObj.getFullYear()}`;
        
        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet) {
            sendSystemAlert(`⚠️ Tab *${targetSheetName}* tidak ditemukan di Google Sheets.`);
            return null;
        }

        return XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Selalu kembalikan Raw Data untuk di-Cache
    } catch (error) {
        sendSystemAlert(`🚨 Error API Google Drive: ${error.message}`);
        return null;
    }
}

function prosesDataKePesanWA(rawData, tanggalAngka = "", keywordCari = "") {
    if (!rawData || !Array.isArray(rawData)) return [];
    let daftarPesanWA = [];
    let blocks = [];
    let currentBlock = [];

    // Kelompokkan baris berdasarkan blok tanggal
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
        // Filter Tanggal
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

            // Filter Pencarian Lanjutan (Smart Search)
            if (keywordCari !== "") {
                if (!msg.toLowerCase().includes(keywordCari.toLowerCase())) {
                    continue; // Jika tidak mengandung keyword pencarian, lewati event ini
                }
            }

            daftarPesanWA.push(msg);
        }
    }
    return daftarPesanWA;
}

function ekstrakStateEvent(rawData) { /* Sama persis, dirangkum untuk efisiensi baris */
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
            let eventKey = `${dateStr}_${c}_${namaTampil}`;
            let crewList = [];
            let statusEvent = "";
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
                }
            }
            state[eventKey] = { nama: namaTampil, tanggal: dateStr, crew: crewList, status: statusEvent, hash: isiEventLengkap.join("~") };
        }
    }
    return state;
}

function cariPerubahanEvent(dataLama, dataBaru) {
    let stateLama = ekstrakStateEvent(dataLama);
    let stateBaru = ekstrakStateEvent(dataBaru);
    let hasilPerubahan = [];
    for (let key in stateBaru) {
        let baru = stateBaru[key];
        let lama = stateLama[key];
        if (!lama || lama.hash !== baru.hash) {
            let msg = `📌 *${baru.nama}* (${baru.tanggal})`;
            msg += baru.crew.length > 0 ? `\n👥 *Crew:* ${baru.crew.join(", ")}` : `\n👥 *Crew:* (Belum diplot)`;
            if (baru.status && baru.status !== "-") msg += `\n🏷️ *Status:* ${baru.status.toUpperCase()}`;
            hasilPerubahan.push(msg);
        }
    }
    return hasilPerubahan;
}


// ============================================================================
// 5. ENGINE WHATSAPP & BOT LOGIC
// ============================================================================
const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"] },
});

// Fitur: Alerting System (Health Check)
const sendSystemAlert = async (text) => {
    console.log(text);
    if (isBotReady) {
        try {
            await client.sendMessage(ID_TUJUAN_NOTIFIKASI, text);
        } catch (e) {}
    }
};

const simulateTyping = async (chat, text) => {
    try {
        await chat.sendSeen();
        await chat.sendStateTyping();
        let typingTime = Math.min(text.length * 30 + 500, 2000);
        await new Promise((resolve) => setTimeout(resolve, typingTime));
        await chat.clearState();
    } catch (error) {}
};

// Fitur Patroli
const jalankanRonda = async () => {
    console.log("🕵️ Meronda...");
    try {
        const dataTerbaru = await getJadwalDariExcel(new Date());
        if (!dataTerbaru) return; // Jika gagal narik excel, stop ronde ini

        if (objekDataLama && JSON.stringify(dataTerbaru) !== JSON.stringify(objekDataLama)) {
            console.log("🔔 Ada perubahan jadwal!");
            const daftarRevisi = cariPerubahanEvent(objekDataLama, dataTerbaru);
            if (daftarRevisi.length > 0) {
                let teksDaftar = daftarRevisi.map((item) => `• *${item}*`).join("\n");
                const pesanNotif = `🚨 *ALARM REVISI ADMIN* 🚨\n\nAdmin baru saja mengubah data pada event:\n${teksDaftar}\n\n💡 _Ketik *1* atau *2* untuk melihat detail peralatan terbaru._`;
                await client.sendMessage(ID_TUJUAN_NOTIFIKASI, pesanNotif);
            }
        }
        objekDataLama = dataTerbaru; // Update Cache Global
    } catch (err) {
        sendSystemAlert(`❌ Sistem Ronda Gagal: ${err.message}`);
    }
};

// --- WHATSAPP EVENT LISTENERS ---
client.on("qr", (qr) => qrcode.generate(qr, { small: true }));

client.on("ready", async () => {
    console.log("✅ Bot Siap!");
    isBotReady = true;

    // Tarik data pertama kali agar In-Memory Cache terisi
    objekDataLama = await getJadwalDariExcel(new Date());

    // 1. Patroli data tiap menit
    setInterval(jalankanRonda, WAKTU_RONDA_MS);

    // 2. Fitur Enterprise: Daily Morning Briefing
    // Terjadwal jalan setiap jam 06:00 pagi setiap hari ('0 6 * * *')
    cron.schedule('0 6 * * *', async () => {
        try {
            const dateObj = new Date();
            const tglHariIni = dateObj.getDate().toString();
            
            // Ambil fresh data khusus pagi hari
            const freshData = await getJadwalDariExcel(dateObj); 
            if(freshData) objekDataLama = freshData;

            const daftarPesan = prosesDataKePesanWA(objekDataLama, tglHariIni, "");
            
            let sapaanPagi = `🌅 *MORNING BRIEFING*\nSelamat pagi, Bli Ari! Hari ini ada *${daftarPesan.length} Event* yang tercatat di sistem.`;
            await client.sendMessage(ID_TUJUAN_NOTIFIKASI, sapaanPagi);

            for (const pesan of daftarPesan) {
                await new Promise((res) => setTimeout(res, 1000));
                await client.sendMessage(ID_TUJUAN_NOTIFIKASI, pesan);
            }
        } catch (error) {
            sendSystemAlert(`❌ Gagal mengirim Morning Briefing: ${error.message}`);
        }
    });

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
    if (msg.from !== ID_TUJUAN_NOTIFIKASI) return;

    const text = msg.body.toLowerCase().trim();
    const chat = await msg.getChat();

    // Menu Utama
    if (["halo", "menu", "jadwal", "bot"].includes(text)) {
        const balasanMenu = `━━━━━━━━━━━━━━━━━━\n🤖 *AGS ENTERPRISE BOT*\n━━━━━━━━━━━━━━━━━━\n\n1️⃣ 📍 Jadwal Hari Ini\n2️⃣ 📍 Jadwal Besok\n3️⃣ 📆 Semua Jadwal Bulan Ini\n\n🔍 *Pencarian Cerdas:*\nKetik \`cari [kata kunci]\`\n_Contoh: cari Bli Ari_\n_Contoh: cari BCA_\n\n✏️ Ketik pilihan Anda...`;
        await simulateTyping(chat, balasanMenu);
        await msg.reply(balasanMenu);
    } 
    
    // Fitur Enterprise: Advanced Query (Pencarian Cerdas)
    else if (text.startsWith("cari ") || text.startsWith("search ")) {
        const keyword = text.replace("cari ", "").replace("search ", "").trim();
        if (keyword.length < 3) {
            return msg.reply("⚠️ Kata kunci terlalu pendek. Minimal 3 huruf.");
        }

        const balasanTunggu = `⏳ Mencari data mengandung kata: *"${keyword}"* ...`;
        await simulateTyping(chat, balasanTunggu);
        await msg.reply(balasanTunggu);

        // Langsung hajar dari CACHE! Super cepat (0ms)
        const dataBulanIni = objekDataLama || (await getJadwalDariExcel(new Date()));
        const daftarPesan = prosesDataKePesanWA(dataBulanIni, "", keyword);

        if (daftarPesan.length === 0) {
            await msg.reply(`ℹ️ Tidak ditemukan event/crew/alat dengan kata kunci *"${keyword}"* bulan ini.`);
        } else {
            await msg.reply(`✅ Ditemukan *${daftarPesan.length} Hasil* pencarian:`);
            for (const pesan of daftarPesan) {
                await simulateTyping(chat, pesan);
                await client.sendMessage(msg.from, pesan);
                await new Promise((res) => setTimeout(res, 500)); 
            }
        }
    }

    // Tarik Data Jadwal (Fitur Enterprise: In-Memory Caching - Kecepatan Kilat)
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
            labelTarget = "Semua Jadwal Bulan Ini";
        }

        const balasanTunggu = `⚡ Menarik data ${labelTarget} (Cache Mode)...`;
        await simulateTyping(chat, balasanTunggu);
        await msg.reply(balasanTunggu);

        // Jika cache kosong karena sistem baru nyala, fetch baru. Kalau sudah ada, pakai cache!
        if (!objekDataLama) {
            objekDataLama = await getJadwalDariExcel(dateObj);
        }
        
        // Proses datanya secara instan dari RAM Server
        const daftarPesan = prosesDataKePesanWA(objekDataLama, tglTarget, "");

        if (daftarPesan.length === 0) {
            const balasanKosong = `ℹ️ Tidak ada jadwal untuk ${labelTarget}.`;
            await simulateTyping(chat, balasanKosong);
            await msg.reply(balasanKosong);
        } else {
            for (const pesan of daftarPesan) {
                await simulateTyping(chat, pesan);
                await client.sendMessage(msg.from, pesan);
                await new Promise((res) => setTimeout(res, 500)); 
            }
        }
    }
});

client.initialize();