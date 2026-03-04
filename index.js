const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const { google } = require('googleapis');
const fs = require('fs').promises;
const readline = require('readline');
const XLSX = require('xlsx');

// --- KONFIGURASI ---
const SPREADSHEET_ID = '1DLcMkga8UiRtRJ3ZQIPMRQb-5d1IFiu_'; // real
// const SPREADSHEET_ID = '18-wJoQ6yLvz17cK0vyNuKyfDs6dhT-8M';
const ID_TUJUAN_NOTIFIKASI = '628970282769@c.us'; 

let memoriDataLama = ""; 

// --- SISTEM LOGIN OAUTH 2.0 ---
async function authorize() {
    let content;
    try {
        content = await fs.readFile('credentials.json');
    } catch (err) {
        console.error('❌ Error: File credentials.json tidak ditemukan!');
        return null;
    }
    const credentials = JSON.parse(content);
    const {client_secret, client_id, redirect_uris} = credentials.installed || credentials.web;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    try {
        const token = await fs.readFile('token.json');
        oAuth2Client.setCredentials(JSON.parse(token));
        return oAuth2Client;
    } catch (err) {
        return await getNewToken(oAuth2Client);
    }
}

async function getNewToken(oAuth2Client) {
    const authUrl = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: ['https://www.googleapis.com/auth/drive.readonly'] });
    console.log('\n=========================================\nBuka link ini:\n' + authUrl + '\n=========================================\n');
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    return new Promise((resolve) => {
        rl.question('Paste kode di sini: ', async (code) => {
            rl.close();
            const {tokens} = await oAuth2Client.getToken(code);
            oAuth2Client.setCredentials(tokens);
            await fs.writeFile('token.json', JSON.stringify(tokens));
            resolve(oAuth2Client);
        });
    });
}

// --- FUNGSI AMBIL DATA EXCEL ---
async function getJadwalDariExcel(tanggalAngka = "", teksTanggal = "", targetDateObj = new Date()) {
    const authClient = await authorize();
    const drive = google.drive({ version: 'v3', auth: authClient });

    try {
        const res = await drive.files.get({ fileId: SPREADSHEET_ID, alt: 'media' }, { responseType: 'arraybuffer' });
        const workbook = XLSX.read(res.data, { type: 'buffer' });
        
        const namaBulan = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
        const targetSheetName = `${namaBulan[targetDateObj.getMonth()]} ${targetDateObj.getFullYear()}`;
        
        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet) return [`❌ Tab *${targetSheetName}* tidak ditemukan.`];

        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (teksTanggal === "RONDA") return rows; 

        return prosesDataKePesanWA(rows, tanggalAngka, teksTanggal);
    } catch (error) {
        console.error(error);
        return ["❌ Gagal mengunduh file Excel."];
    }
}

// --- FUNGSI PEMBANTU: KONVERSI TANGGAL EXCEL KE TEKS ---
function formatTanggalExcel(val) {
    if (!val) return "-";
    
    // Jika bentuknya angka (Excel Serial Date), konversi ke format Indonesia
    if (!isNaN(val) && val > 40000) {
        const date = new Date(Math.round((val - 25569) * 86400 * 1000));
        const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        return `${date.getDate()} ${namaBulan[date.getMonth()]} ${date.getFullYear()}`;
    }
    
    return val.toString().trim();
}

// --- FUNGSI PARSING DATA (FIX: QTY + NAMA ALAT + FREQ) ---
function prosesDataKePesanWA(rawData, tanggalAngka = "", teksTanggal = "") {
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
            const getVal = (r, col) => (block[r] && block[r][col] ? block[r][col].toString().trim() : "");
            
            if (getVal(1, c).toUpperCase() !== "NAME") continue; 
            
            const eventTitle = getVal(2, c + 6);
            if (!eventTitle || eventTitle === "-" || eventTitle === "Event Tittle" || eventTitle === "") continue; 

            const customerName = getVal(1, c + 1) || "-";
            const companyName  = getVal(2, c + 1) || "-";
            const dateEventRaw = block[1] ? block[1][c+6] : "-"; 
            const dateEvent    = formatTanggalExcel(dateEventRaw);
            const venue        = getVal(3, c + 6) || "-";
            const loadingDate  = getVal(4, c + 6) || "-";
            
            let crewList = [];
            let itemList = [];
            
            for (let i = 8; i < block.length; i++) {
                const marker = getVal(i, c).toUpperCase();
                if (marker.includes("STATUS") || marker.includes("CUSTOMER")) break;
                
                const crew = getVal(i, c + 6);
                if (crew && crew !== "-" && crew !== "CREW" && crew !== "") crewList.push(crew);
                
                // --- PENENTUAN KOLOM EXCEL ---
                const item = getVal(i, c + 1); // ITEM
                const spec = getVal(i, c + 2); // SPESIFICATION
                const qty  = getVal(i, c + 3); // QTY
                const freq = getVal(i, c + 5); // FREQUENCY (Berapa hari)

                if (qty && item && item.toUpperCase() !== "ITEM") {
                    // Hasilnya: "* 2 Projector 5000 Lumen (1 hari)"
                    let teksAlat = `* ${qty} ${item} ${spec}`;
                    if (freq && freq !== "-") {
                        teksAlat += ` (${freq})`; // Tambahkan durasi jika ada
                    }
                    itemList.push(teksAlat.replace(/\s+/g, ' ').trim());
                }
            }

            // let msg = `📢 *JADWAL EVENT [TGL ${masterTanggal}]* 📢\n\n`;
            // msg += `🏢 *Klien:* ${companyName} (${customerName})\n`;
            // msg += `🎪 *Event:* ${eventTitle}\n`;
            // msg += `📍 *Venue:* ${venue}\n`;
            // msg += `📅 *Tgl:* ${dateEvent}\n`;
            // msg += `⏰ *Loading:* ${loadingDate}\n\n`;
            // msg += `👥 *Crew:* ${crewList.length > 0 ? crewList.join(', ') : '-'}\n\n`;
            // msg += `🎛️ *Daftar Alat:*\n${itemList.length > 0 ? itemList.join('\n') : '- Data alat kosong -'}`;

            // --- IDE TEMPLATE PESAN LEBIH RAPI ---
            let msg = `━━━━━ 📝 *DETAIL EVENT* ━━━━━\n\n`;

            // SEKTOR INFO UTAMA
            msg += `📌 *EVENT:* ${eventTitle}\n`;
            msg += `🏢 *KLIEN:* ${companyName}\n`;
            msg += `📍 *VENUE:* ${venue}\n`;
            msg += `📅 *TANGGAL:* ${dateEvent}\n`;
            msg += `🚚 *LOADING:* ${loadingDate}\n\n`;

            // SEKTOR TIM
            msg += `👥 *TIM BERTUGAS (CREW):*\n`;
            msg += crewList.length > 0 ? crewList.map(c => `   ◦ ${c}`).join('\n') : `   - (Belum ada kru)`;
            msg += `\n\n`;

            // SEKTOR ALAT (Menggunakan bullet point yang konsisten)
            msg += `📦 *DAFTAR ALAT & DURASI:*\n`;
            if (itemList.length > 0) {
                msg += itemList.join('\n');
            } else {
                msg += `   - (Data alat kosong)`;
            }

            msg += `\n\n━━━━━━━━━━━━━━━━━━━━`;
            
            daftarPesanWA.push(msg);
        }
    }
    return daftarPesanWA;
}

// --- WHATSAPP CLIENT ---
const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: { args: ['--no-sandbox', '--disable-setuid-sandbox'] }
});

client.on('qr', (qr) => qrcode.generate(qr, { small: true }));

client.on('ready', async () => {
    console.log('✅ Bot Siap!');
    const dataAwal = await getJadwalDariExcel("", "RONDA", new Date());
    memoriDataLama = JSON.stringify(dataAwal);

    setInterval(async () => {
        console.log('🕵️ Meronda...');
        const dataTerbaruRaw = await getJadwalDariExcel("", "RONDA", new Date());
        const dataTerbaruStr = JSON.stringify(dataTerbaruRaw);
        if (memoriDataLama !== "" && dataTerbaruStr !== memoriDataLama) {
            await client.sendMessage(ID_TUJUAN_NOTIFIKASI, `🚨 *ALARM REVISI ADMIN* 🚨\n\nAda perubahan data di Excel! Ketik *1* atau *2* untuk cek.`);
            memoriDataLama = dataTerbaruStr;
        }
    }, 10 * 60 * 1000);
});

client.on('message', async (msg) => {
    const text = msg.body.toLowerCase().trim();

    if (['halo', 'menu', 'jadwal', 'bot'].includes(text)) {
        await msg.reply(`🤖 *MENU JADWAL*\n\n1️⃣ Hari Ini\n2️⃣ Besok\n3️⃣ Semua Jadwal Bulan Ini`);
    } 
    else if (['1', '2', '3'].includes(text)) {
        const date = new Date();
        let tgl = text === '1' ? date.getDate().toString() : (text === '2' ? (date.setDate(date.getDate() + 1), date.getDate().toString()) : "");
        let label = text === '1' ? "Hari Ini" : (text === '2' ? "Besok" : "Semua Jadwal Bulan Ini");

        await msg.reply(`⏳ Menarik data ${label}...`);
        const daftarPesan = await getJadwalDariExcel(tgl, label, date);
        
        // PENGIRIMAN PESAN BERTAHAP (Anti-Crash)
        if (daftarPesan.length === 0 || typeof daftarPesan === 'string') {
            await msg.reply(typeof daftarPesan === 'string' ? daftarPesan : `ℹ️ Tidak ada jadwal untuk ${label}.`);
        } else {
            for (const pesan of daftarPesan) {
                await client.sendMessage(msg.from, pesan);
                // Beri jeda 1 detik tiap pesan agar tidak kena spam block
                await new Promise(res => setTimeout(res, 1000));
            }
        }
    }
});

client.initialize();