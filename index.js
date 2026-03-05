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
    if (!isNaN(val) && val > 40000) {
        const date = new Date(Math.round((val - 25569) * 86400 * 1000));
        const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        return `${date.getDate()} ${namaBulan[date.getMonth()]} ${date.getFullYear()}`;
    }
    return val.toString().trim();
}

// --- FUNGSI PARSING DATA ---
// --- FUNGSI PEMBANTU: PENYORTIR KATEGORI ALAT ---
// --- FUNGSI PEMBANTU: PENYORTIR KATEGORI ALAT (VERSI UPDATE INVENTORY) ---
function tentukanKategori(namaAlat) {
    const teks = namaAlat.toLowerCase();

    const kamusKategori = {
        "📺 VISUAL & MULTIMEDIA": [
            'videotron', 'tv', 'monitor', 'projector', 'screen', 'kamera', 
            'switcher', 'klicker', 'perfect cue', 'laptop', 'timer'
        ],
        "💡 LIGHTING": [
            'moving', 'strobe', 'fresnel', 'par led', 'avolite', 'grandma', 'grand ma',
            'lighting', 'beam', 'smoke', 'hazer', 'efx', 'minuit', 'tripod t'
        ],
        "🔊 SOUND & BACKLINE": [
            'console', 'speaker', 'subwoofer', 'mic', 'yamaha', 'midas', 
            'dl32', 'foh', 'mixer', 'in ear', 'stand mic', 'audio focus', 
            'iem', 'drumset', 'tama', 'sound system', 'milan', 'sp milan', 
            'pa ', 'senheiser', 'sennheiser', 'roland', 'akustika'
        ],
        "🏗️ RIGGING & STAGING": [
            'rigging', 'rig', 'gawangan', 'level', 'aluminium', 'stage', 
            'barikade', 'tenda'
        ],
        "⚡ POWER": [
            'genset', 'kabel', 'power', 'panel', 'distro'
        ]
    };

    for (const [kategori, kataKunciArray] of Object.entries(kamusKategori)) {
        const cocok = kataKunciArray.some(kataKunci => teks.includes(kataKunci));
        if (cocok) {
            return kategori;
        }
    }

    return "📦 LAINNYA"; 
}

// --- FUNGSI PARSING DATA (VERSI TEMPLATE PREMIUM) ---
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

            const companyName  = getVal(2, c + 1) || "-";
            const dateEventRaw = block[1] ? block[1][c+6] : "-"; 
            const dateEvent    = formatTanggalExcel(dateEventRaw);
            const venue        = getVal(3, c + 6) || "-";
            const loadingDate  = getVal(4, c + 6) || "-";
            
            let crewList = [];
            
            // --- INI YANG HARUS DIPERBAIKI ---
            // Wadah ini namanya HARUS SAMA PERSIS dengan hasil dari tentukanKategori()
            let kategoriAlat = {
                "📺 VISUAL & MULTIMEDIA": [],
                "💡 LIGHTING": [],
                "🔊 SOUND & BACKLINE": [],
                "🏗️ RIGGING & STAGING": [],
                "⚡ POWER": [],
                "📦 LAINNYA": []
            };
            
            for (let i = 8; i < block.length; i++) {
                const marker = getVal(i, c).toUpperCase();
                if (marker.includes("STATUS") || marker.includes("CUSTOMER")) break;
                
                const crew = getVal(i, c + 6);
                if (crew && crew !== "-" && crew !== "CREW" && crew !== "") crewList.push(crew);
                
                const item = getVal(i, c + 1); 
                const spec = getVal(i, c + 2); 
                const qty  = getVal(i, c + 3); 
                const freq = getVal(i, c + 5); 

                if (qty && item && item.toUpperCase() !== "ITEM") {
                    let namaLengkap = `${item} ${spec}`.trim();
                    let teksAlat = `• ${qty} ${namaLengkap}`;
                    
                    if (freq && freq !== "-") teksAlat += ` (${freq})`;
                    
                    teksAlat = teksAlat.replace(/\s+/g, ' ').trim();
                    
                    // Proses sortir
                    let namaKategori = tentukanKategori(namaLengkap);
                    
                    // Pastikan kategori ada sebelum di-push (Mencegah error undefined)
                    if (kategoriAlat[namaKategori]) {
                        kategoriAlat[namaKategori].push(teksAlat);
                    } else {
                        kategoriAlat["📦 LAINNYA"].push(teksAlat);
                    }
                }
            }

            // --- MULAI MENYUSUN PESAN SESUAI TEMPLATE ---
            let msg = `━━━━━━━━━━━━━━━━━━━━\n📝 *EVENT DETAIL*\n━━━━━━━━━━━━━━━━━━━━\n\n`;
            
            msg += `📌 *EVENT* : ${eventTitle}\n`;
            msg += `🏢 *CLIENT* : ${companyName}\n`;
            msg += `📍 *VENUE* : ${venue}\n`;
            msg += `📅 *DATE* : ${dateEvent}\n`;
            msg += `🚚 *LOADING*: ${loadingDate}\n\n`;
            
            msg += `━━━━━━━━━━━━━━━━━━━━\n👥 *CREW*\n`;
            msg += crewList.length > 0 ? crewList.map(cr => `• ${cr}`).join('\n') : `• (Belum ada crew)`;
            msg += `\n\n`;

            // Loop untuk menampilkan kategori alat hanya jika ada isinya
            for (const [namaKat, listKat] of Object.entries(kategoriAlat)) {
                if (listKat.length > 0) {
                    msg += `━━━━━━━━━━━━━━━━━━━━\n${namaKat}\n`;
                    msg += listKat.join('\n') + `\n\n`;
                }
            }

            msg = msg.trim() + `\n━━━━━━━━━━━━━━━━━━━━`;
            
            daftarPesanWA.push(msg);
        }
    }
    return daftarPesanWA;
}

// --- WHATSAPP CLIENT DENGAN KONFIGURASI SERVER ---
const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: { 
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-background-timer-throttling',
            '--disable-dev-shm-usage',
        ]
    }
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

// --- FUNGSI SIMULASI NGETIK ---
const simulateTyping = async (chat, text) => {
    // Tandai pesan sudah dibaca (opsional, agar centang biru muncul duluan)
    await chat.sendSeen(); 
    
    await chat.sendStateTyping();
    // Hitung durasi: 50 milidetik per karakter + jeda mikir 500 milidetik
    // Maksimal delay dibatasi 3 detik agar orang tidak menunggu kelamaan
    let typingTime = (text.length * 50) + 500;
    if (typingTime > 3000) typingTime = 3000; 
    
    await new Promise(resolve => setTimeout(resolve, typingTime));
};

// --- SISTEM ANTI-ZOMBIE (AUTO RESTART) ---

// 1. Jika WhatsApp memutus koneksi
client.on('disconnected', (reason) => {
    console.log('❌ Bot terputus dari WhatsApp! Alasan:', reason);
    console.log('🔄 Mematikan proses agar direstart ulang oleh PM2...');
    process.exit(1); // Ini akan memaksa Node.js mati, dan PM2 akan otomatis menghidupkannya lagi
});

// 2. Jika ada error tak terduga (agar script tidak nge-hang)
process.on('unhandledRejection', (error) => {
    console.error('⚠️ Ada error tak tertangkap:', error.message);
    // Jangan di-exit di sini, cukup log saja agar bot tetap jalan
});

// --- WHATSAPP MESSAGE HANDLER ---
client.on('message', async (msg) => {
    const text = msg.body.toLowerCase().trim();
    
    // Tarik data chat untuk keperluan efek ngetik dan auto-read
    const chat = await msg.getChat();

    if (['halo', 'menu', 'jadwal', 'bot'].includes(text)) {
        const balasanMenu = `🤖 *MENU JADWAL*\n\n1️⃣ Hari Ini\n2️⃣ Besok\n3️⃣ Semua Jadwal Bulan Ini`;
        
        await simulateTyping(chat, balasanMenu);
        await msg.reply(balasanMenu);
    } 
    else if (['1', '2', '3'].includes(text)) {
        const date = new Date();
        let tgl = text === '1' ? date.getDate().toString() : (text === '2' ? (date.setDate(date.getDate() + 1), date.getDate().toString()) : "");
        let label = text === '1' ? "Hari Ini" : (text === '2' ? "Besok" : "Semua Jadwal Bulan Ini");

        const balasanTunggu = `⏳ Menarik data ${label}...`;
        await simulateTyping(chat, balasanTunggu);
        await msg.reply(balasanTunggu);
        
        const daftarPesan = await getJadwalDariExcel(tgl, label, date);
        
        if (daftarPesan.length === 0 || typeof daftarPesan === 'string') {
            const balasanKosong = typeof daftarPesan === 'string' ? daftarPesan : `ℹ️ Tidak ada jadwal untuk ${label}.`;
            await simulateTyping(chat, balasanKosong);
            await msg.reply(balasanKosong);
        } else {
            for (const pesan of daftarPesan) {
                await simulateTyping(chat, pesan);
                await client.sendMessage(msg.from, pesan);
                await new Promise(res => setTimeout(res, 1000));
            }
        }
    }
});

client.initialize();