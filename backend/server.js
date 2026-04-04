import 'dotenv/config';
import { removePlaceholderText, paraphrase } from "./utils/util.js";
import prisma from './utils/prisma.js';
import express from 'express';
import cors from 'cors';
import { GoogleGenerativeAI } from '@google/generative-ai';


const app = express();
const port = process.env.PORT || 8000;
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const GOOGLE_CLOUD_PROJECT = process.env.GOOGLE_CLOUD_PROJECT;
const GOOGLE_CLOUD_LOCATION = process.env.GOOGLE_CLOUD_LOCATION || 'global';
console.log(GOOGLE_CLOUD_LOCATION)
// let iterasi = 0
if (!GEMINI_API_KEY) {
    console.log('Gemini API KEY tidak ditemukan');
    process.exit(1);
}

const genAI = new GoogleGenerativeAI(GEMINI_API_KEY);


app.use(cors({
    origin: "https://paraphraser-add-in.vercel.app",
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS', 'PATCH'],
    // allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin']
}))

app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', req.headers.origin || '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS, PATCH');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');

    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
})

app.use(express.json({ limit: '50mb' }))
app.use(express.urlencoded({ extended: true }))

// Request timeout middleware (60 detik)
const REQUEST_TIMEOUT_MS = 60 * 1000;
app.use((req, res, next) => {
    req.setTimeout(REQUEST_TIMEOUT_MS);
    res.setTimeout(REQUEST_TIMEOUT_MS);

    req.on('timeout', () => {
        if (!res.headersSent) {
            console.error(`⏱️ Request timeout (${REQUEST_TIMEOUT_MS / 1000}s): ${req.method} ${req.originalUrl}`);
            res.status(408).json({ error: 'Request timeout: proses melebihi batas waktu 60 detik' });
        }
    });

    next();
});

// Endpoint dummy untuk simulasi hasil gagal (testing frontend)
app.get('/api/paraphrase/failed-test', (req, res) => {
    const failedResults = {
        paraphrasedText: [
            "Pemanfaatan AI pada penyusunan karya akademis mampu menaikkan efisiensi waktu pengerjaan skripsi untuk mahasiswa tahun terakhir.",
            "Aplikasi kecerdasan buatan di penulisan karya ilmiah sangat menolong percepatan penyelesaian tugas akhir bagi mahasiswa semester akhir.",
            "Implementasi teknologi AI dalam menulis karya ilmiah bisa mendukung efisiensi waktu penyelesaian tugas akhir untuk mahasiswa tingkat akhir."
        ],
        paraphrased: [
            "Pemanfaatan AI pada penyusunan karya akademis mampu menaikkan efisiensi waktu pengerjaan skripsi untuk mahasiswa tahun terakhir.",
            "Aplikasi kecerdasan buatan di penulisan karya ilmiah sangat menolong percepatan penyelesaian tugas akhir bagi mahasiswa semester akhir.",
            "Implementasi teknologi AI dalam menulis karya ilmiah bisa mendukung efisiensi waktu penyelesaian tugas akhir untuk mahasiswa tingkat akhir."
        ],
        pScore: [-11, 0.78, 0.92],
        bScore: [0.1, 0.55, 0.75],
        logId: "dummy-test-log-id"
    };

    console.log('🧪 Mengirim dummy failedResult ke frontend');
    res.json(failedResults);
});

app.post('/api/paraphrase', async (req, res) => {
    try {
        const { originalText, tokenizedText, objects } = req.body;
        let iterasi = 0;

        // Objek untuk menampung hasil jika 3x percobaan tetap tidak memenuhi threshold
        let failedResults = {
            paraphrasedText: [],
            paraphrased: [],
            pScore: [],
            bScore: [],
        };
        console.log("ORIGINAL TEXT: ", originalText)
        console.log("TOKENIZED: ", tokenizedText)
        console.log("OBJECTS", objects)
        let sourceText = tokenizedText;
        sourceText = removePlaceholderText(sourceText);

        if (!sourceText || typeof sourceText !== 'string') {
            return res.status(400).json({ error: 'sourceText wajib dan harus string' });
        }
        console.log("source text --------------------------", sourceText)

        // ── Buat ParaphraseLog SEBELUM loop (1 log per request pengguna) ──────
        const logEntry = await prisma.paraphraseLog.create({
            data: { isSuccessful: false }  // update setelah loop selesai
        });
        const logId = logEntry.id;
        console.log('📝 ParaphraseLog created, id:', logId);

        let successResult = null;
        let paraPLUIE_score = -1;
        let BLEU_score = 1;

        // ── LOOP: tiap iterasi memanggil paraphrase() yang auto-save ke tabel Paraphrase ──
        while ((paraPLUIE_score < 0 || BLEU_score > 0.7) && iterasi < 3) {
            iterasi++;
            let paraphraseResult = await paraphrase(
                sourceText,
                objects,
                logId,          // paraphraseLogId  → paraphrase() simpan ke DB
                originalText,   // originalText
                tokenizedText,  // tokenizedText
                iterasi         // iterationNumber (1-based)
            );

            if (typeof paraphraseResult === "string") {
                console.log("ERROR INVALID TEXT")
                await prisma.paraphraseLog.update({
                    where: { id: logId },
                    data: { isSuccessful: false, errorLog: 'Teks tidak valid', iterations: iterasi }
                });
                return res.status(400).json({ error: 'Teks yang dimasukkan tidak valid' });
            }

            paraPLUIE_score = paraphraseResult.pScore;
            BLEU_score = paraphraseResult.bScore;

            if (paraPLUIE_score >= 0 && BLEU_score <= 0.7) {
                successResult = paraphraseResult;
                break;
            } else {
                failedResults.paraphrasedText.push(paraphraseResult.paraphrasedText);
                failedResults.paraphrased.push(paraphraseResult.paraphrased);
                failedResults.pScore.push(paraPLUIE_score);
                failedResults.bScore.push(BLEU_score);
            }
        }
        console.log("iterasi = ", iterasi)

        // ── Update ParaphraseLog dengan outcome akhir ─────────────────────────
        if (successResult) {
            await prisma.paraphraseLog.update({
                where: { id: logId },
                data: { isSuccessful: true, iterations: iterasi }
            });
            console.log('✅ ParaphraseLog updated (success), id:', logId);
            res.json({ ...successResult, logId });
            console.log("sukses:", successResult)
        } else {
            await prisma.paraphraseLog.update({
                where: { id: logId },
                data: { isSuccessful: false, iterations: iterasi }
            });
            console.log('⚠️ ParaphraseLog updated (failed threshold), id:', logId);
            res.json({ ...failedResults, logId });
            console.log("failresult:", failedResults)
        }

    } catch (e) {
        console.error('Error:', e);
        try {
            await prisma.paraphraseLog.create({
                data: {
                    isSuccessful: false,
                    errorLog: e.message || String(e),
                }
            });
        } catch (dbErr) {
            console.error('DB logging error (catch):', dbErr);
        }
        res.status(500).json({ error: e.message || 'Gagal memparafrase teks' });
    }
});

// Endpoint untuk menyimpan feedback (rating & komentar) ke database
app.patch('/api/paraphrase/:id/feedback', async (req, res) => {
    try {
        const { id } = req.params;
        const { rating, userComment } = req.body;

        if (!rating || typeof rating !== 'number' || rating < 1 || rating > 5) {
            return res.status(400).json({ error: 'Rating harus angka 1-5' });
        }

        const updated = await prisma.paraphraseLog.update({
            where: { id },
            data: {
                rating: rating,
                userComment: userComment || null,
            }
        });

        console.log('✅ Feedback saved for log id:', id);
        res.json({ success: true, id: updated.id });
    } catch (e) {
        console.error('Feedback error:', e);
        
        if (e.code === 'P2025') {
            return res.status(404).json({ error: 'Log tidak ditemukan' });
        }
        res.status(500).json({ error: e.message || 'Gagal menyimpan feedback' });
    }
});

// Endpoint untuk menyimpan error dari frontend ke database
app.post('/api/log-error', async (req, res) => {
    try {
        const { errorMessage } = req.body;

        const logEntry = await prisma.paraphraseLog.create({
            data: {
                isSuccessful: false,
                errorLog: errorMessage || 'Unknown frontend error',
            }
        });

        console.log('⚠️ Frontend error logged to database, id:', logEntry.id);
        res.json({ success: true, id: logEntry.id });
    } catch (e) {
        console.error('Error logging failed:', e);
        res.status(500).json({ error: e.message || 'Gagal menyimpan error log' });
    }
});

app.listen(port, '0.0.0.0', () => {
    console.log(`Backend server running on port ${port}`);
    console.log(`Intelligem API endpoint: http://localhost:${port}/api/intelligem`);
});
