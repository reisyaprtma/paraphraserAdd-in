/* * Taskpane.js
 * Mengatur interaksi UI, Dark Mode, dan Simulasi Logika sesuai Proposal Bab VI
 */

/**
 * Fungsi inti untuk memanggil API backend Intelligem.
 * @param {string} originalText Prompt teks dari pengguna.
 * @param {string} selectedMode
 * @param {Array} object Teks yang disorot di dokumen.
*/

// ----------------------------------------- GLOBAL VARIABEL --------------------------
// const BACKEND_URL = "http://localhost:8000/api/paraphrase";
const BACKEND_URL = "https://paraphraseradd-in-779882211224.asia-southeast3.run.app/api/paraphrase";

// let resultPara = ""
import { htmlToText } from 'html-to-text';
let finalOoxml = ""
let finalParaphrased
let paraphrasedRaw
let tokenizedTexts
let objects
let seleksi_teks
let isHidden = true
let finalOoxml_array = [];
let currentLogId = null; // ID dari ParaphraseLog di database

// --- XML BOILERPLATE CONSTANTS ---
const xmlHeader = `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" 
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <w:body>`;
const xmlFooter = `
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

//------------------------------------------ GET SELECTION AND STARTER ----------------------------
/** Memperbarui isi kotak teks dengan teks yang sedang diseleksi di dokumen. */
async function updateSelectionInBox() {
    const box = document.getElementById('source-text');
    const statusEl = document.getElementById('selection-status');
    const btn = document.getElementById("paraphrase-btn");
    // const btn2 = document.getElementById("paraphrase-fast-btn");
    if (!box) return;
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();

            const text = selection.text || '';
            const segmenter = new Intl.Segmenter('en', { granularity: 'word' });
            const segments = segmenter.segment(text);

            // Use an iterator to avoid memory issues with huge strings
            let count = 0;
            for (const segment of segments) {
                if (segment.isWordLike) count++;
            }
            console.log(count); // Output: 6


            if (count > 750) {
                btn.disabled = true;
                btn.classList.add("opacity-50");
                // btn2.disabled = true;
                // btn2.classList.add("opacity-50");
            }
            else if (count === 0) {
                btn.disabled = true;
                btn.classList.add("opacity-50");
                // btn2.disabled = true;
                // btn2.classList.add("opacity-50");
            }
            else {
                btn.disabled = false;
                btn.classList.remove("opacity-50");
                // btn2.disabled = false;
                // btn2.classList.remove("opacity-50");
            }
            box.value = text;
            if (statusEl) {
                if (count > 750) {
                    statusEl.innerHTML = `<i data-lucide="alert-triangle" class="text-red-500 w-3 h-3"></i> Jumlah kata melebihi batas maksimal! (Terseleksi ${count} kata)`;
                    // statusEl.classList.add('active');
                    statusEl.classList.remove('active');
                    statusEl.classList.add('bg-red-100');
                    statusEl.classList.add('text-red-500');
                } else if (count > 0) {
                    statusEl.innerHTML = `<i data-lucide="check" class="w-3 h-3"></i> ${count} kata terseleksi`;
                    statusEl.classList.add('active');
                } else {
                    statusEl.innerHTML = `<i data-lucide="mouse-pointer-2" class="w-3 h-3"></i> Tidak ada seleksi`;
                    statusEl.classList.remove('active');
                    statusEl.classList.remove('bg-red-100');
                    statusEl.classList.remove('text-red-500');
                }
                if (typeof lucide !== 'undefined') lucide.createIcons();
            }
        });
    } catch (e) {
        console.warn('Gagal baca seleksi:', e);
        if (statusEl) {
            statusEl.innerHTML = `<i data-lucide="mouse-pointer-2" class="w-3 h-3"></i> Menunggu seleksi…`;
            statusEl.classList.remove('active');
        }
    }
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("selection-status").innerHTML = `<i data-lucide="mouse-pointer-2" class="w-3 h-3"></i> Tidak ada teks terseleksi`;
        setupEventListeners();

        // Auto-run test untuk melihat tampilan showFailedResult


        // Isi awal dari seleksi saat add-in dibuka
        updateSelectionInBox();

        // Perbarui real time saat seleksi berubah di dokumen
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            () => updateSelectionInBox(),
            (err) => {
                if (err) console.warn('DocumentSelectionChanged handler gagal:', err);
            }
        );
    }
});

//------------------------------------------ PREPROCESS FUNCTION --------------------------
async function extractCombinedTokens() {
    let mappingData = {
        math: [],
        citations: []
    };
    let finalTokenizedText = "";
    
    // ❌ let seleksi_teks = ""; <-- BARIS INI SUDAH DIHAPUS AGAR MENGGUNAKAN GLOBAL VARIABLE

    await Word.run(async (context) => {
        // 1️⃣ SETUP & LOAD AWAL
        const selection = context.document.getSelection();
        const contentControls = selection.contentControls;
        
        selection.load("text");
        contentControls.load("items/text, items/tag");
        
        const mainOoxmlResult = selection.getOoxml();
        await context.sync();

        // ✅ SEKARANG INI AKAN MENGISI VARIABEL GLOBAL
        seleksi_teks = selection.text;
        console.log("Original Text:", seleksi_teks);

        // ---------------------------------------------------------
        // MULAI PEMROSESAN DATA
        // ---------------------------------------------------------
        let currentXml = mainOoxmlResult.value || "";
        if (!currentXml) return;

        let citationIndex = 0;

        // 2️⃣ PASS 1: MENDELEY LAMA & ZOTERO TOKENIZATION (Menggunakan RegEx)
        // RegEx diperlebar untuk memakan seluruh blok Field dari "begin" sampai "end"
        // Mendeley lama : ADDIN CSL_CITATION
        // Zotero        : ADDIN ZOTERO_ITEM CSL_CITATION
        const cslRegex = /<w:fldChar[^>]*w:fldCharType="begin"[^>]*>[\s\S]*?ADDIN (?:ZOTERO_ITEM )?CSL_CITATION[\s\S]*?<w:fldChar[^>]*w:fldCharType="end"[^>]*>/gi;
        
        currentXml = currentXml.replace(cslRegex, (match) => {
            const token = `[[CIT_${citationIndex + 1}]]`;
            
            let extractedRawText = "[Referensi]";
            try {
                // Zotero menggunakan "plainCitation" dan "formattedCitation"
                // Mendeley menggunakan "plainTextFormattedCitation" dan "formattedCitation"
                const rawTextMatch = 
                    match.match(/"plainTextFormattedCitation":"(.*?)"/) ||
                    match.match(/"plainCitation":"(.*?)"/) ||
                    match.match(/"formattedCitation":"(.*?)"/);
                if (rawTextMatch && rawTextMatch[1]) {
                    extractedRawText = rawTextMatch[1];
                }
            } catch (error) {
                console.warn("Gagal mengekstrak raw text:", error);
            }
            
            mappingData.citations.push({
                token: token,
                ooxml: match, // Sekarang ini menyimpan blok Field yang UTUH sempurna
                raw: extractedRawText,
                isRegex: true 
            });
            citationIndex++;
            return token; // Menggantikan seluruh blok (termasuk [4]) dengan [[CIT_n]]
        });
        console.log("currentXml: ============================================================\n", currentXml)

        // 3️⃣ PASS 2: MATH TOKENIZATION
        let mathIndex = 0;
        const oMathRegex = /<m:oMath[\s\S]*?<\/m:oMath>|<m:oMathPara[\s\S]*?<\/m:oMathPara>/gi;

        currentXml = currentXml.replace(oMathRegex, (match) => {
            const token = ` [[MATH_${mathIndex + 1}]] `;
            let rawMath = match.replace(/<[^>]+>/g, "").trim();

            mappingData.math.push({
                token: token.trim(),
                ooxml: match,
                raw: rawMath 
            });
            mathIndex++;
            return token;
        });

        // 4️⃣ PASS 3: CLEANING XML TO TEXT
        // Jika tidak ada sitasi DAN tidak ada math yang terdeteksi,
        // gunakan teks plain dari selection.text secara langsung.
        // Ini mencegah kebocoran teks mentah dari field citation manager lain
        // (misal: EndNote, RefWorks) yang tidak dikenali dan tidak tertangkap PASS 1.
        // Jika ada math → tetap wajib proses dari XML agar posisi [[MATH_n]] benar.
        let textBase;
        if (citationIndex === 0 && mathIndex === 0) {
            console.log("⚡ Tidak ada sitasi maupun math → menggunakan plain text langsung (bypass XML cleaning)");
            textBase = seleksi_teks;
        } else {
            textBase = currentXml
                .replace(/<\/w:p>/gi, "\n")
                .replace(/<w:br\/>/gi, "\n")
                .replace(/<w:tab\/>/gi, "\t")
                .replace(/<[^>]+>/g, "")
                .replace(/&lt;/g, "<")
                .replace(/&gt;/g, ">")
                .replace(/&amp;/g, "&")
                .replace(/&quot;/g, '"')
                .replace(/&apos;/g, "'")
                .replace(/[ \t]+/g, " ").trim();
        }

        // 5️⃣ PASS 4: MENDELEY BARU TOKENIZATION (Content Controls)
        const citationRequests = contentControls.items.map((cc) => {
            return {
                text: cc.text, 
                ooxmlProxy: cc.getOoxml()
            };
        });

        await context.sync(); // Wajib sync lagi karena getOoxml() bikin proxy baru

        citationRequests.forEach((req) => {
            const citToken = `[[CIT_${citationIndex + 1}]]`;

            mappingData.citations.push({
                token: citToken,
                ooxml: req.ooxmlProxy.value, 
                raw: req.text,
                isRegex: false
            });

            // Ganti teks asli dengan token
            textBase = textBase.split(req.text).join(citToken);
            citationIndex++;
        });

        finalTokenizedText = textBase;
    });

    return {
        raw_text: seleksi_teks,
        text: finalTokenizedText,
        object: mappingData
    };
}
// Helper: Log error ke database via backend
async function logErrorToDatabase(errorMessage) {
    try {
        const ERROR_URL = BACKEND_URL.replace('/api/paraphrase', '/api/log-error');
        await fetch(ERROR_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                sourceText: seleksi_teks || '',
                tokenizedInput: tokenizedTexts || null,
                errorMessage: errorMessage,
            })
        });
        console.log('Error logged to database');
    } catch (logErr) {
        console.error('Gagal log error ke database:', logErr);
    }
}

//------------------------------------------ BACKEND CALLING --------------------------
async function callAPI(originalText, tokenizedText, objects, mode) {
    try {
        const payload = {
            originalText: originalText,
            tokenizedText: tokenizedText,
            objects: objects,
            mode: mode,
        }
        const response = await fetch(BACKEND_URL, {
            method: 'POST',
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(payload)
        })
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            const message = errorData.error || response.statusText || `HTTP ${response.status}`;
            throw new Error(message);
        }

        const data = await response.json()
        const paraphrasedText = data.paraphrasedText
        finalParaphrased = paraphrasedText
        paraphrasedRaw = data.paraphrased
        const paraphrased = data.paraphrased
        const pScore = data.pScore
        const bScore = data.bScore
        currentLogId = data.logId || null; // Simpan ID untuk feedback nanti
        if (typeof paraphrasedText == 'string') {
            showResult(paraphrasedText, paraphrased, pScore, bScore)
        }
        else if (Array.isArray(paraphrased)) {
            showFailedResult(paraphrasedText, paraphrased, pScore, bScore)
        }
        // const result = JSON.parse(data.paraphraseText)
        // console.log(result)

        // replaceToken()
        console.log("Data yang diterima:", data)

    } catch (error) {
        console.error("Kesalahan saat memanggil Intelligem API:", error.message);
        showError(`${error.message}`);
        await logErrorToDatabase(error.message);
    }
}

async function sendPrompt(mode) {
    const diffToggle = document.getElementById("diff-toggle");
    const toggleCircle2 = document.getElementById("toggle-circle-2");
    const failedresultSection = document.getElementById("failed-results-container");
    const resultSection = document.getElementById("result-section");
    const warningBox = document.getElementById("warning-box");
    const errorBox = document.getElementById("error-box");

    // UI Reset: Sembunyikan semua container hasil sebelumnya
    if (failedresultSection) failedresultSection.classList.add("!hidden");
    if (resultSection) resultSection.classList.add("hidden");
    if (warningBox) warningBox.classList.add('hidden');
    if (errorBox) errorBox.classList.add('hidden');

    if (diffToggle) diffToggle.classList.remove("active");
    if (toggleCircle2) toggleCircle2.classList.remove("translate-x-5");

    const starBtns = document.querySelectorAll('.star-btn');
    const submitBtn = document.getElementById('submit-feedback-btn');
    const comment = document.getElementById('feedback-comment');
    comment.value = "";
    comment.disabled = false;
    // Reset semua dimensi rating
    currentRating = { semantic: 0, gramatical: 0, syntactic: 0, lexical: 0 };
    starBtns.forEach(btn => {
        btn.disabled = false;
        btn.classList.remove('opacity-50');
        btn.classList.remove('text-yellow-400');

        // Reset SVG icon color back to default gray (unselected)
        const icon = btn.querySelector('svg');
        if (icon) {
            icon.classList.remove('text-yellow-400', 'fill-yellow-400');
            icon.classList.add('text-gray-300');
        }
    });
    submitBtn.disabled = false;
    submitBtn.classList.remove('opacity-50');
    submitBtn.innerText = "Kirim Penilaian";
    submitBtn.classList.remove('bg-green-100', 'text-green-700', 'dark:bg-green-900', 'dark:text-green-300');

    const extractedToken = await extractCombinedTokens()
    tokenizedTexts = extractedToken.text
    objects = extractedToken.object
    console.log(tokenizedTexts)
    console.log(objects)
    const tokenizedText = tokenizedTexts
    try {
        const loadingState = document.getElementById("loading-state");
        const btn = document.getElementById("paraphrase-btn");
        // const btn2 = document.getElementById("paraphrase-fast-btn");
        // const selectedMode = document.getElementById("paraphrase-mode")
        // UI Transition: Start Loading
        btn.disabled = true;
        btn.classList.add("opacity-50");
        // btn2.disabled = true;
        // btn2.classList.add("opacity-50");
        loadingState.classList.remove("hidden");
        // resultSection.classList.add("hidden"); // Sudah disembunyikan di awal
        console.log(seleksi_teks)
        await callAPI(seleksi_teks, tokenizedText, objects, mode)

        // console.log('berhasil')
    } catch (error) {
        console.error("Kesalahan saat mengirim prompt ke Intelligem:", error);
        showError(`Terjadi kesalahan: ${error.message}. Periksa console.`);
        await logErrorToDatabase(error.message);
    }
}

//------------------------------------------ POST PROCCESS ---------------------
function ooxmlResult(resultPara, objects) {
    if (!resultPara || !objects) {
        showError("Belum ada hasil parafrase atau data token hilang.");
        return;
    }

    const tokenRegex = /(\[\[(?:CIT|MATH)_\d+\]\])/g;
    let newOoxmlBody = "";

    // Split baris (handle Enter)
    const lines = resultPara.split(/\r?\n/);

    lines.forEach(line => {
        // Skip baris kosong/sampah
        if (!line.trim() || line.includes("Click or tap here")) return;

        newOoxmlBody += "<w:p>"; // Buka Paragraf Baru

        const parts = line.split(tokenRegex);

        parts.forEach(part => {
            if (!part) return;

            if (tokenRegex.test(part)) {
                let mappingFound = null;
                let isMath = part.includes("MATH");
                let isCit = part.includes("CIT");

                if (isCit) {
                    mappingFound = objects.citations.find(item => item.token === part);
                } else if (isMath) {
                    mappingFound = objects.math.find(item => item.token === part);
                }

                if (mappingFound && mappingFound.ooxml) {
                    let contentToInsert = "";

                    if (isMath) {
                        // KASUS MATH: Gunakan langsung raw XML
                        contentToInsert = mappingFound.ooxml;

                        // 🔥 FIXED LOGIC MATH: JANGAN Rename, tapi UNWRAP.
                        // Math Para biasanya bentuknya: <m:oMathPara><m:oMath>...</m:oMath></m:oMathPara>
                        // Kita hanya perlu buang kulit <m:oMathPara>-nya, agar tersisa <m:oMath> (inline)
                        contentToInsert = contentToInsert
                            .replace(/<m:oMathPara[\s\S]*?>/g, "") // Hapus pembuka Para
                            .replace(/<\/m:oMathPara>/g, "");      // Hapus penutup Para

                        // PENTING: Jika math aslinya Inline (<m:oMath>), regex di atas tidak match,
                        // jadi isinya tetap <m:oMath> (Aman, tidak double).
                    }
                    else if (isCit) {
                        if (mappingFound.isRegex) {
                            // Cukup bungkus kembali dengan <w:r> untuk menjaga keseimbangan tag XML
                            contentToInsert = `<w:r>${mappingFound.ooxml}</w:r>`; 
                        } else {
                            // Jika dari Mendeley Baru (Content Controls), ekstrak body-nya seperti biasa
                            contentToInsert = extractBodyFromOoxml(mappingFound.ooxml);
                            contentToInsert = contentToInsert
                                .replace(/<w:p [^>]*>/g, "")
                                .replace(/<w:p>/g, "")
                                .replace(/<\/w:p>/g, "")
                                .replace(/<w:sectPr[^>]*>[\s\S]*?<\/w:sectPr>/g, "");
                        }
                    }

                    newOoxmlBody += contentToInsert;

                } else {
                    // Fallback Merah
                    newOoxmlBody += `<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>${escapeXml(part)}</w:t></w:r>`;
                }
            } else {
                // Teks Biasa
                newOoxmlBody += `<w:r><w:t xml:space="preserve">${escapeXml(part)}</w:t></w:r>`;
            }
        });

        newOoxmlBody += "</w:p>"; // Tutup Paragraf
    });

    const finalOoxml = xmlHeader + newOoxmlBody + xmlFooter;

    console.log(finalOoxml)

    return finalOoxml
}

/**
 * Mengambil isi di dalam tag <w:body> dari string OOXML lengkap.
 * Ini penting karena getOoxml() mengembalikan paket lengkap, 
 * sedangkan kita hanya butuh isinya untuk digabung.
 */
function extractBodyFromOoxml(fullOoxml) {
    // Regex mencari konten di antara <w:body> ... </w:body>
    // Flag 's' (dotAll) tidak selalu support di semua browser lama, jadi pakai [\s\S]*
    const bodyRegex = /<w:body[^>]*>([\s\S]*?)<\/w:body>/i;
    const match = fullOoxml.match(bodyRegex);

    if (match && match[1]) {
        return match[1]; // Kembalikan isinya saja
    }
    return ""; // Gagal ekstrak
}

/**
 * Membersihkan karakter yang dilarang dalam XML (<, >, &, dll)
 */
function escapeXml(unsafe) {
    if (!unsafe) return "";
    return unsafe.replace(/[<>&'"]/g, function (c) {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\'': return '&apos;';
            case '"': return '&quot;';
        }
    });
}

//------------------------------------------ SHOW RESULT AND INSERT ---------------------
// ── Sentence Tracker State ──
let sentenceSpans = [];
let currentSentenceIndex = -1; // -1 = show all
const activeSentenceIndices = {};

function renderStackedDiff(original, paraphrased) {
    const diffs = Diff.diffWords(original, paraphrased);

    let unifiedHtml = "";
    let remCount = 0;
    let addCount = 0;

    // Buffer untuk menampung "Cluster" perubahan
    let bufferRem = "";
    let bufferAdd = "";

    // Fungsi Helper: Menentukan apakah sebuah teks adalah "Jangkar" (Kata tetap yang valid)
    // Kita anggap teks itu "Jangkar" jika mengandung huruf/angka.
    // Kalau cuma spasi atau tanda baca (misal ", ."), kita anggap itu "lem" yang bisa ikut masuk ke kotak perubahan.
    const isAnchor = (text) => /[a-zA-Z0-9]/.test(text);

    // Fungsi untuk merender apa yang ada di buffer
    const flushBuffer = () => {
        // Render Merah (Dihapus) Full satu blok
        // Render Hijau (Ditambah) Full satu blok
        if (bufferRem.trim().length > 0) {
            unifiedHtml += `<span class="removal-highlight hidden bg-red-500/5 text-gray-800 dark:text-gray-200 line-through decoration-red-500">${bufferRem}</span>`;
        }
        if (bufferAdd.trim().length > 0) {
            unifiedHtml += `<span class="addition-highlight bg-green-200 dark:bg-green-800 text-green-800 dark:text-green-200">${bufferAdd}</span>`;

        }
        // Reset Buffer
        bufferRem = "";
        bufferAdd = "";
    };

    diffs.forEach((part) => {
        // KASUS 1: Bagian yang BERUBAH (Added/Removed)
        if (part.removed) {
            bufferRem += part.value;
            remCount++;
        }
        else if (part.added) {
            bufferAdd += part.value;
            addCount++;
        }
        // KASUS 2: Bagian TETAP (Unchanged)
        else {
            // Cek: Apakah ini kata beneran atau cuma spasi/koma?
            if (isAnchor(part.value)) {
                // Kalau ketemu KATA BENERAN (Jangkar), baru kita flush buffer sebelumnya.
                // Ini yang bikin efek "Ketika Anda mencari" (Merah semua) vs "Saat menelusuri" (Hijau semua)
                flushBuffer();

                // Render kata tetapnya
                unifiedHtml += `<span>${part.value}</span>`;
            } else {
                // Kalau cuma spasi/tanda baca, JANGAN flush dulu.
                // Masukkan spasi ini ke KEDUA buffer agar kalimatnya nanti nyambung/rapi di dalam kotak.
                // (Kecuali jika buffer masih kosong, berarti ini spasi awal kalimat biasa)
                if (bufferRem.length > 0 || bufferAdd.length > 0) {
                    bufferRem += part.value;
                    bufferAdd += part.value;
                } else {
                    unifiedHtml += `<span>${part.value}</span>`;
                }
            }
        }
    });

    // Flush sisa buffer di akhir (jika kalimat diakhiri dengan perubahan)
    flushBuffer();

    // Output ke DOM (tanpa modifikasi HTML structure)
    document.getElementById("result-content").innerHTML = unifiedHtml;

    // ── Sentence Tracking: Post-process DOM ──
    setupSentenceTracking(paraphrased);

    const diffToggle = document.getElementById("diff-toggle");
    const toggleCircle2 = document.getElementById("toggle-circle-2");
    // 1. Add the dot for the class selector
    const removedTexts = document.querySelectorAll(".removal-highlight");

    if (diffToggle) {
        diffToggle.onclick = () => {

            // 2. We use a variable to track state based on the first element (if it exists)

            // 3. Loop through every highlighted span and toggle it
            // 4. Move the circle based on the state
            removedTexts.forEach((el) => {
                isHidden = el.classList.toggle("hidden");
            });
            toggleCircle2.classList.toggle(isHidden);
            toggleCircle2.classList.toggle("translate-x-5")
            diffToggle.classList.toggle("active");
            console.log("isHidden on render: ", isHidden)

            // if (isHidden) {
            //     diffToggle.classList.remove("bg-indigo-600");
            //     toggleCircle2.classList.remove("translate-x-5")
            // } else {
            //     diffToggle.classList.add("bg-indigo-600");
            //     toggleCircle2.classList.add("translate-x-5")
            // }
            // diffToggle.classList.remove("bg-indigo-600");
            // toggleCircle2.classList.remove("translate-x-5")

        };
    }
    // Update Counter
    // if (document.getElementById("removal-count")) document.getElementById("removal-count").innerText = remCount;
    // if (document.getElementById("addition-count")) document.getElementById("addition-count").innerText = addCount;
    const segmenter = new Intl.Segmenter('id', { granularity: 'word' });
    const segments = segmenter.segment(paraphrased);

    // Use an iterator to avoid memory issues with huge strings
    let count = 0;
    for (const segment of segments) {
        if (segment.isWordLike) count++;
    }
    console.log("jumlah kata: ", count); // Output: 6
    document.getElementById("addition-count").innerText = `${count} Kata`;   
}

/**
 * Cari posisi karakter dimana kalimat baru dimulai
 * Berdasarkan tanda baca akhir kalimat (. ? !) yang diikuti spasi
 */
function findSentenceBoundaries(text) {
    const boundaries = [0]; // Kalimat pertama mulai di 0
    const regex = /[.!?]\s+/g;
    let match;
    
    while ((match = regex.exec(text)) !== null) {
        const boundaryPos = match.index + match[0].length;
        if (boundaryPos < text.length) {
            boundaries.push(boundaryPos);
        }
    }
    
    return boundaries;
}

/**
 * POST-PROCESSING APPROACH: Setelah diff HTML dirender ke DOM,
 * walk semua child span, track posisi teks visible (skip removal-highlight),
 * dan assign data-sentence attribute ke setiap span.
 * Jika suatu span memuat boundary kalimat, split span tersebut
 * menjadi 2+ span via DOM cloneNode (HTML selalu valid).
 */
function setupSentenceTracking(paraphrasedText, contentId = 'result-content', positionId = 'sentence-position', prevBtnId = 'sentence-prev', nextBtnId = 'sentence-next') {
    const container = document.getElementById(contentId);
    if (!container) return;
    
    // Hitung sentence boundaries dari teks parafrase
    const boundaries = findSentenceBoundaries(paraphrasedText);
    const totalSentences = boundaries.length;
    
    // Ambil semua child element saat ini
    const children = Array.from(container.children);
    
    if (children.length === 0) {
        initSentenceTracker(contentId, positionId, prevBtnId, nextBtnId);
        return;
    }
    
    // Jika cuma 1 kalimat, mark semua sebagai sentence 0
    if (totalSentences <= 1) {
        children.forEach(child => {
            child.setAttribute('data-sentence', '0');
            child.classList.add('sentence-span');
        });
        initSentenceTracker(contentId, positionId, prevBtnId, nextBtnId);
        return;
    }
    
    // Walk children, track posisi teks visible, assign sentence
    let visiblePos = 0;
    let curSentence = 0;
    
    children.forEach(child => {
        if (child.nodeType !== Node.ELEMENT_NODE) return;
        
        const isRemoved = child.classList.contains('removal-highlight');
        
        if (isRemoved) {
            // Teks dihapus: assign ke kalimat saat ini, JANGAN advance posisi
            child.setAttribute('data-sentence', String(curSentence));
            child.classList.add('sentence-span');
            return;
        }
        
        // Untuk teks visible (unchanged/addition): cek apakah ada boundary crossing
        const text = child.textContent || '';
        const startPos = visiblePos;
        const endPos = visiblePos + text.length;
        
        // Cari semua boundaries yang jatuh di dalam rentang span ini
        const crossings = [];
        for (let b = 0; b < boundaries.length; b++) {
            if (boundaries[b] > startPos && boundaries[b] < endPos) {
                crossings.push({
                    offset: boundaries[b] - startPos,  // offset relatif dalam teks span
                    sentenceIdx: b
                });
            }
        }
        
        if (crossings.length === 0) {
            // Tidak ada boundary crossing, simple
            child.setAttribute('data-sentence', String(curSentence));
            child.classList.add('sentence-span');
        } else {
            // Perlu split span ini di titik-titik boundary
            const fragment = document.createDocumentFragment();
            let lastOffset = 0;
            let segSentence = curSentence;
            
            crossings.forEach(({ offset, sentenceIdx }) => {
                const segText = text.substring(lastOffset, offset);
                if (segText.length > 0) {
                    const newSpan = child.cloneNode(false); // clone tanpa children
                    newSpan.textContent = segText;
                    newSpan.setAttribute('data-sentence', String(segSentence));
                    newSpan.classList.add('sentence-span');
                    fragment.appendChild(newSpan);
                }
                segSentence = sentenceIdx;
                lastOffset = offset;
            });
            
            // Sisa teks setelah boundary terakhir
            const remaining = text.substring(lastOffset);
            if (remaining.length > 0) {
                const newSpan = child.cloneNode(false);
                newSpan.textContent = remaining;
                newSpan.setAttribute('data-sentence', String(segSentence));
                newSpan.classList.add('sentence-span');
                fragment.appendChild(newSpan);
            }
            
            child.replaceWith(fragment);
            curSentence = segSentence;
        }
        
        // Advance posisi visible
        visiblePos = endPos;
        
        // Update curSentence berdasarkan posisi terkini
        while (curSentence < boundaries.length - 1 && visiblePos >= boundaries[curSentence + 1]) {
            curSentence++;
        }
    });
    
    initSentenceTracker(contentId, positionId, prevBtnId, nextBtnId);
}

/**
 * Inisialisasi sentence tracker UI
 * Menggunakan data-sentence attribute, bukan .sentence-span wrapping
 */
function initSentenceTracker(contentId = 'result-content', positionId = 'sentence-position', prevBtnId = 'sentence-prev', nextBtnId = 'sentence-next') {
    const container = document.getElementById(contentId);
    if (!container) return;
    
    // Hitung unique sentences dari data-sentence attributes
    const allMarked = container.querySelectorAll('[data-sentence]');
    const sentenceSet = new Set();
    allMarked.forEach(el => sentenceSet.add(el.getAttribute('data-sentence')));
    const totalSentences = sentenceSet.size;
    
    activeSentenceIndices[contentId] = -1; // Reset: tampilkan semua
    
    const posEl = document.getElementById(positionId);
    const prevBtn = document.getElementById(prevBtnId);
    const nextBtn = document.getElementById(nextBtnId);
    
    if (posEl) posEl.textContent = "Semua kalimat";
    
    // Add click handler pada setiap sentence-span
    allMarked.forEach(el => {
        el.addEventListener('click', () => {
            const idx = parseInt(el.getAttribute('data-sentence'));
            if (!isNaN(idx)) navigateToSentence(idx, contentId, positionId, prevBtnId, nextBtnId);
        });
    });
    
    // Button handlers
    if (prevBtn) {
        prevBtn.onclick = () => {
            const currentIndex = activeSentenceIndices[contentId] ?? -1;
            // Jika sedang di posisi paling awal (0), navigasi prev -> Semua Kalimat (-1)
            if (currentIndex >= 0) {
                navigateToSentence(currentIndex - 1, contentId, positionId, prevBtnId, nextBtnId);
            }
        };
    }
    
    if (nextBtn) {
        nextBtn.onclick = () => {
            const currentIndex = activeSentenceIndices[contentId] ?? -1;
            // Jika mencapai posisi maksimum (terakhir), navigasi next -> Semua Kalimat (-1)
            if (currentIndex < totalSentences - 1) {
                const nextIdx = currentIndex === -1 ? 0 : currentIndex + 1;
                navigateToSentence(nextIdx, contentId, positionId, prevBtnId, nextBtnId);
            } else {
                showAllSentences(contentId, positionId, prevBtnId, nextBtnId);
            }
        };
    }
    
    updateTrackerButtons(contentId, positionId, prevBtnId, nextBtnId);
}

/**
 * Navigate ke kalimat tertentu
 */
function navigateToSentence(index, contentId = 'result-content', positionId = 'sentence-position', prevBtnId = 'sentence-prev', nextBtnId = 'sentence-next') {
    if (index === -1) {
        showAllSentences(contentId, positionId, prevBtnId, nextBtnId);
        return;
    }
    
    activeSentenceIndices[contentId] = index;
    const container = document.getElementById(contentId);
    if (!container) return;
    
    const allMarked = container.querySelectorAll('[data-sentence]');
    let firstActive = null;
    
    allMarked.forEach(el => {
        const elSentence = parseInt(el.getAttribute('data-sentence'));
        el.classList.remove('sentence-active', 'sentence-dimmed');
        
        if (elSentence === index) {
            el.classList.add('sentence-active');
            if (!firstActive) firstActive = el;
        } else {
            el.classList.add('sentence-dimmed');
        }
    });
    
    // Scroll ke kalimat aktif pertama
    if (firstActive) {
        firstActive.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    
    updateTrackerButtons(contentId, positionId, prevBtnId, nextBtnId);
}

/**
 * Tampilkan semua kalimat (reset)
 */
function showAllSentences(contentId = 'result-content', positionId = 'sentence-position', prevBtnId = 'sentence-prev', nextBtnId = 'sentence-next') {
    activeSentenceIndices[contentId] = -1;
    const container = document.getElementById(contentId);
    if (!container) return;
    
    const allMarked = container.querySelectorAll('[data-sentence]');
    allMarked.forEach(el => {
        el.classList.remove('sentence-active', 'sentence-dimmed');
    });
    
    updateTrackerButtons(contentId, positionId, prevBtnId, nextBtnId);
}

/**
 * Update state tombol navigasi
 */
function updateTrackerButtons(contentId = 'result-content', positionId = 'sentence-position', prevBtnId = 'sentence-prev', nextBtnId = 'sentence-next') {
    const container = document.getElementById(contentId);
    const posEl = document.getElementById(positionId);
    const prevBtn = document.getElementById(prevBtnId);
    const nextBtn = document.getElementById(nextBtnId);
    
    if (!posEl || !container) return;
    
    // Hitung total sentences
    const sentenceSet = new Set();
    container.querySelectorAll('[data-sentence]').forEach(el => {
        sentenceSet.add(el.getAttribute('data-sentence'));
    });
    const totalSentences = sentenceSet.size;
    
    const currentIndex = activeSentenceIndices[contentId] ?? -1;
    
    if (currentIndex === -1) {
        // Mode "semua"
        posEl.innerHTML = `Semua kalimat <span class="capitalize lowercase ml-1" style="opacity:0.6; font-size:10px">(${totalSentences})</span>`;
        if (prevBtn) prevBtn.disabled = true;
        // Next bisa di-klik selama jumlah kalimat lebih dari 1
        if (nextBtn) nextBtn.disabled = totalSentences <= 1;
    } else {
        posEl.textContent = `Kalimat ${currentIndex + 1} dari ${totalSentences}`;
        // Prev selalu aktif kapanpun asal index tidak lebih kecil dari 0 (jika 0 maka lari ke Semua)
        if (prevBtn) prevBtn.disabled = false; 
        if (nextBtn) nextBtn.disabled = false; // Next selalu aktif, jika di kalimat terakhir lari ke "Semua Kalimat"
    }
}

async function showResult(tokenizedResult, data, parapluie, bleu) {
    const loadingState = document.getElementById("loading-state");
    const resultSection = document.getElementById("result-section");
    const btn = document.getElementById("paraphrase-btn");
    // const btn2 = document.getElementById("paraphrase-fast-btn");

    loadingState.classList.add("hidden");
    resultSection.classList.remove("hidden");
    btn.disabled = false;
    btn.classList.remove("opacity-50");
    // btn2.disabled = false;
    // btn2.classList.remove("opacity-50");

    console.log("Displayed Text to User:", data);

    renderStackedDiff(seleksi_teks, data);
    finalOoxml = ooxmlResult(tokenizedResult, objects);

    // [BARU] Update Bar Visualisasi
    // updateHealthCheckUI(parapluie, bleu);
    // [BARU] Re-render lucide icon setelah display diubah
    lucide.createIcons();
}

// seleksi_teks = "Penggunaan teknologi kecerdasan buatan dalam penulisan karya ilmiah dapat membantu meningkatkan efisiensi waktu pengerjaan tugas akhir bagi mahasiswa tingkat akhir.";
function showFailedResult(tokenizedResultArray, dataArray, parapluieArray, bleuArray) {
    const loadingState = document.getElementById("loading-state");
    const failedresultSection = document.getElementById("failed-results-container");
    const warningBox = document.getElementById("warning-box");
    // const starRating = document.getElementById("star-rating");
    // starRating.classList.toggle('hidden')
    warningBox.classList.remove('hidden')
    const btn = document.getElementById("paraphrase-btn");
    // const btn2 = document.getElementById("paraphrase-fast-btn");
    let mappingData = {
        math: [],
        citations: []
    };

    // 1. Reset & Show UI
    loadingState.classList.add("hidden");
    failedresultSection.classList.remove("!hidden", "hidden");
    btn.disabled = false;
    btn.classList.remove("opacity-50");
    // btn2.disabled = false;
    // btn2.classList.remove("opacity-50");

    // 2. Process OOXML for EACH result in the array
    // Assuming tokenizedResultArray is an array of token sets corresponding to dataArray
    finalOoxml_array = tokenizedResultArray.map((tokens) => {
        return ooxmlResult(tokens, objects);
    });

    console.log("Generated OOXML Array:", finalOoxml_array);


    // FOR TEST --------------------------------------------------------------------------------
    // seleksi_teks = `
    // Activity diagram ini menjelaskan alur aktivitas Admin dalam mengelola data pekerjaan pada sistem informasi. Proses diawali ketika Admin melakukan login ke dalam sistem. Setelah berhasil masuk, Admin mengakses tab Manajemen Pekerjaan untuk melihat data pekerjaan yang sudah tersedia. Sistem kemudian akan menampilkan data pekerjaan yang tersimpan di dalam basis data
    // `
    // 3. Render all results dynamically
    // We pass the ooxml array here so the renderer can attach it to the buttons
    failedRenderStackedDiff(seleksi_teks, dataArray, parapluieArray, bleuArray, finalOoxml_array);
}

function failedRenderStackedDiff(original, paraphrasedArray, pScores, bScores, ooxmlData) {
    // Gunakan container khusus untuk hasil gagal (bisa menampung banyak kartu)
    const container = document.getElementById("failed-results-container");
    if (!container) return;

    // Bersihkan hasil sebelumnya
    container.innerHTML = "";

    // Untuk setiap hasil parafrase yang gagal, buat satu kartu mirip UI sukses
    paraphrasedArray.forEach((paraphrasedText, index) => {
        const diffs = Diff.diffWords(original.trim(), paraphrasedText.trim());

        let unifiedHtml = "";
        let remCount = 0;
        let addCount = 0;

        // Logika clustering sama seperti renderStackedDiff (sukses)
        let bufferRem = "";
        let bufferAdd = "";
        const isAnchor = (text) => /[a-zA-Z0-9]/.test(text);

        const flushBuffer = () => {
            if (bufferRem.trim().length > 0) {
                unifiedHtml += `<span class="removal-highlight hidden bg-red-500/5 text-gray-800 dark:text-gray-200 line-through decoration-red-500">${bufferRem}</span>`;
            }
            if (bufferAdd.trim().length > 0) {
                unifiedHtml += `<span class="addition-highlight bg-green-200 dark:bg-green-800 text-green-800 dark:text-green-200">${bufferAdd}</span>`;
            }
            bufferRem = "";
            bufferAdd = "";
        };

        diffs.forEach((part) => {
            if (part.removed) {
                bufferRem += part.value;
                remCount++;
            } else if (part.added) {
                bufferAdd += part.value;
                addCount++;
            } else {
                if (isAnchor(part.value)) {
                    flushBuffer();
                    unifiedHtml += `<span>${part.value}</span>`;
                } else {
                    if (bufferRem.length > 0 || bufferAdd.length > 0) {
                        bufferRem += part.value;
                        bufferAdd += part.value;
                    } else {
                        unifiedHtml += `<span>${part.value}</span>`;
                    }
                }
            }
        });

        flushBuffer();

        // Hitung jumlah kata
        const segmenter = new Intl.Segmenter('id', { granularity: 'word' });
        const segments = segmenter.segment(paraphrasedText);
        let wordCount = 0;
        for (const segment of segments) {
            if (segment.isWordLike) wordCount++;
        }

        const card = document.createElement("div");
        card.className = "card animate-slide-up mb-4";

        card.innerHTML = `
            <!-- Card Header -->
            <div class="card-header flex justify-between items-center">
                <div class="flex items-center gap-2">
                    <span class="opsi-badge">Opsi ${index + 1}</span>
                    <span class="section-label" style="margin-bottom:0">Tampilkan Teks Dibuang</span>
                </div>
                <button id="diff-toggle-failed-${index}" class="diff-toggle-btn">
                    <div id="toggle-circle-failed-${index}" class="diff-toggle-thumb"></div>
                </button>
            </div>

            <!-- Diff Content -->
            <div class="card-body">
                <div id="result-content-failed-${index}">
                    ${unifiedHtml}
                </div>
            </div>

            <!-- Sentence Tracker Minimalis -->
            <div id="sentence-tracker-failed-${index}" class="flex items-center justify-between px-4 py-2 border-b" style="border-color: var(--clr-border); background: rgba(0,0,0,0.015);">
                <button id="sentence-prev-failed-${index}" class="btn-icon" style="width:24px;height:24px" title="Kalimat Sebelumnya" disabled>
                    <i data-lucide="chevron-left" class="w-3.5 h-3.5"></i>
                </button>
                <span id="sentence-position-failed-${index}" class="text-[11px] font-semibold tracking-wide text-gray-500 uppercase">Semua kalimat</span>
                <button id="sentence-next-failed-${index}" class="btn-icon" style="width:24px;height:24px" title="Kalimat Berikutnya" disabled>
                    <i data-lucide="chevron-right" class="w-3.5 h-3.5"></i>
                </button>
            </div>

            <!-- Card Footer: counts + action buttons -->
            <div class="card-footer">
                <div class="diff-count">
                    <span class="section-label">Jumlah Kata: <span id="addition-count-failed-${index}">${wordCount} Kata</span></span>
                </div>
                <div class="flex gap-2">
                    <button id="copy-btn-failed-${index}" class="btn-icon" title="Salin Teks">
                        <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                    </button>
                    <button id="insert-btn-failed-${index}" class="btn-icon filled" title="Sisipkan ke Dokumen">
                        <i data-lucide="download" class="w-3.5 h-3.5"></i>
                    </button>
                </div>
            </div>
        `;

        container.appendChild(card);

        // Setup Sentence Tracking secara independen untuk kartu ini
        setupSentenceTracking(
            paraphrasedText,
            `result-content-failed-${index}`,
            `sentence-position-failed-${index}`,
            `sentence-prev-failed-${index}`,
            `sentence-next-failed-${index}`
        );

        // Toggle untuk masing-masing kartu (hanya mempengaruhi removal-highlight di kartu ini)
        const diffToggle = card.querySelector(`#diff-toggle-failed-${index}`);
        const toggleCircle2 = card.querySelector(`#toggle-circle-failed-${index}`);
        const removedTexts = card.querySelectorAll(".removal-highlight");
        let localHidden = true;

        if (diffToggle && toggleCircle2) {
            diffToggle.onclick = () => {
                removedTexts.forEach((el) => {
                    localHidden = el.classList.toggle("hidden");
                });
                toggleCircle2.classList.toggle("translate-x-5");
                diffToggle.classList.toggle("active");
            };
        }

        // Tombol sisipkan khusus untuk opsi ini
        const insertBtn = card.querySelector(`#insert-btn-failed-${index}`);
        if (insertBtn) {
            insertBtn.addEventListener("click", function () {
                if (objects.math.length === 0 && objects.citations.length === 0) {
                    insert(paraphrasedText);
                } else if (ooxmlData && ooxmlData[index]) {
                    insertxml(ooxmlData[index]);
                }
                showNotification("Teks opsi " + (index + 1) + " berhasil disisipkan.");
            });
        }

        // Tombol copy khusus untuk opsi ini
        const copyBtn = card.querySelector(`#copy-btn-failed-${index}`);
        if (copyBtn) {
            copyBtn.addEventListener("click", async function () {
                try {
                    await navigator.clipboard.writeText(paraphrasedText);
                    showNotification("Teks opsi " + (index + 1) + " berhasil disalin.");
                    console.log("Teks opsi " + (index + 1) + " disalin.");
                } catch (err) {
                    console.error("Gagal menyalin", err);
                }
            });
        }
    });

    // Re-render Lucide icons
    if (typeof lucide !== 'undefined') lucide.createIcons();
}


async function insertxml(finalOoxml) {
    console.log(finalOoxml)
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertOoxml(finalOoxml, Word.InsertLocation.replace);
        console.log("OOXML berhasil disisipkan.", finalOoxml);
        await context.sync();
    }).catch(error => {
        console.error("Gagal Insert OOXML:", error);
        // console.log("DEBUG XML BODY:", newOoxmlBody); 
        showError("Gagal menyisipkan teks. Cek Console.");
    });
}
async function insert(text) {
    console.log(text)
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.replace);
        await context.sync();
    }).catch(error => {
        console.error("Gagal Insert Text:", error);
        // console.log("DEBUG XML BODY:", newOoxmlBody); 
        showError("Gagal menyisipkan teks. Cek Console.");
    });
}

function showError(msg) {
    console.error(msg);
    // Bisa disambungkan ke elemen UI error di HTML kamu
    const loadingState = document.getElementById("loading-state");
    const btn = document.getElementById("paraphrase-btn");
    // const btn2 = document.getElementById("paraphrase-fast-btn");
    loadingState.classList.add("hidden");
    btn.disabled = false;
    btn.classList.remove("opacity-50");
    // btn2.disabled = false;
    // btn2.classList.remove("opacity-50");
    const errorBox = document.getElementById('error-box')
    const errorMessage = document.getElementById('error-message')
    errorBox.classList.remove('hidden')
    errorMessage.innerHTML = msg
}

function test() {
    // 1. Mock teks asli yang seolah-olah diseleksi oleh user di Word
    seleksi_teks = "Penggunaan teknologi kecerdasan buatan dalam penulisan karya ilmiah dapat membantu meningkatkan efisiensi waktu pengerjaan tugas akhir bagi mahasiswa tingkat akhir.";

    // 2. Mock global 'objects' agar fungsi ooxmlResult tidak error saat memproses token
    objects = {
        math: [],
        citations: []
    };

    // 3. Mock respons dari backend (Format Array)
    const mockParaphrasedTextArray = [
        "Pemanfaatan AI pada penyusunan karya akademis mampu menaikkan efisiensi waktu pengerjaan skripsi untuk mahasiswa tahun terakhir.",
        "Aplikasi kecerdasan buatan di penulisan karya ilmiah sangat menolong percepatan penyelesaian tugas akhir bagi mahasiswa semester akhir.",
        "Implementasi teknologi AI dalam menulis karya ilmiah bisa mendukung efisiensi waktu penyelesaian tugas akhir untuk mahasiswa tingkat akhir."
    ];

    const mockParaphrasedArray = [
        "Pemanfaatan AI pada penyusunan karya akademis mampu menaikkan efisiensi waktu pengerjaan skripsi untuk mahasiswa tahun terakhir.",
        "Aplikasi kecerdasan buatan di penulisan karya ilmiah sangat menolong percepatan penyelesaian tugas akhir bagi mahasiswa semester akhir.",
        "Implementasi teknologi AI dalam menulis karya ilmiah bisa mendukung efisiensi waktu penyelesaian tugas akhir untuk mahasiswa tingkat akhir."
    ];

    const mockPScoreArray = [-11, 0.78, 0.92];
    const mockBScoreArray = [0.1, 0.55, 0.75];

    console.log("🛠️ Menjalankan simulasi hasil Array...");

    // 4. Panggil fungsi showFailedResult persis seperti yang dilakukan di callAPI
    showFailedResult(mockParaphrasedTextArray, mockParaphrasedArray, mockPScoreArray, mockBScoreArray);

    // 5. Re-render Lucide icons di kartu yang baru dibuat
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

// Fungsi untuk fetch data dummy failedResult dari backend endpoint /api/paraphrase/failed-test
async function fetchFailedTest() {
    const FAILED_TEST_URL = BACKEND_URL.replace('/api/paraphrase', '/api/paraphrase/failed-test');
    const loadingState = document.getElementById("loading-state");
    const failedResultSection = document.getElementById("failed-results-container");
    const btn = document.getElementById("test-btn");
    const resultSection = document.getElementById("result-section");
    const warningBox = document.getElementById("warning-box");
    const errorBox = document.getElementById("error-box");

    // Mock teks asli dan objects untuk rendering diff
    seleksi_teks = "Penggunaan teknologi kecerdasan buatan dalam penulisan karya ilmiah dapat membantu meningkatkan efisiensi waktu pengerjaan tugas akhir bagi mahasiswa tingkat akhir.";
    objects = { math: [], citations: [] };

    // UI: Sembunyikan SEMUA container hasil sebelum mulai loading
    btn.disabled = true;
    btn.classList.add("opacity-50");
    loadingState.classList.remove("hidden");

    if (failedResultSection) failedResultSection.classList.add("!hidden");
    if (resultSection) resultSection.classList.add("hidden");
    if (warningBox) warningBox.classList.add("hidden");
    if (errorBox) errorBox.classList.add("hidden");

    try {
        console.log("🧪 Fetching dummy failed result dari backend...");
        const response = await fetch(FAILED_TEST_URL, {
            method: 'GET',
            headers: { 'Content-Type': 'application/json' }
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();
        console.log("🧪 Data dummy diterima:", data);

        currentLogId = data.logId || null;

        // Tampilkan hasil gagal menggunakan fungsi yang sudah ada
        showFailedResult(data.paraphrasedText, data.paraphrased, data.pScore, data.bScore);

        // Re-render Lucide icons
        if (typeof lucide !== 'undefined') lucide.createIcons();
    } catch (error) {
        console.error("🧪 Gagal fetch failed-test:", error);
        showError(`Gagal mengambil data test: ${error.message}`);
    } finally {
        btn.disabled = false;
        btn.classList.remove("opacity-50");
    }
}

function check() {
    const removal = document.getElementById("removed-container")
    removal.classList.remove('hidden')
}

function setupEventListeners() {
    // 1. Settings & Dark Mode Toggle
    const settingsBtn = document.getElementById("settings-btn");
    const settingsMenu = document.getElementById("settings-menu");
    const helpBtn = document.getElementById("help-btn");
    const helpMenu = document.getElementById("help-menu");
    const darkModeToggle = document.getElementById("dark-mode-toggle");
    const warningBox = document.getElementById("warning-box")
    warningBox.classList.add('hidden')
    const errorBox = document.getElementById("error-box")
    errorBox.classList.add('hidden')
    const htmlEl = document.documentElement;
    const toggleCircle = document.getElementById("toggle-circle");

    settingsBtn.onclick = () => {
        settingsMenu.classList.toggle("hidden");
        helpMenu.classList.add("hidden");
    };

    helpBtn.onclick = () => {
        helpMenu.classList.toggle("hidden");
        settingsMenu.classList.add("hidden");
    };

    darkModeToggle.onclick = () => {
        htmlEl.classList.toggle("dark");
        if (htmlEl.classList.contains("dark")) {
            toggleCircle.classList.add("translate-x-5");
        } else {
            toggleCircle.classList.remove("translate-x-5");
        }
    };

    // 2. Paraphrase Mode Selector
    const modeBtns = document.querySelectorAll('.mode-btn');
    // Set default active mode attribute on the Formal button
    const defaultMode = document.getElementById('mode-formal');
    if (defaultMode) defaultMode.setAttribute('data-active', 'true');

    modeBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            modeBtns.forEach(b => {
                b.removeAttribute('data-active');
                b.style.background = 'transparent';
                b.style.color = 'var(--clr-text, #555)';
                b.style.boxShadow = 'none';
            });
            btn.setAttribute('data-active', 'true');
            btn.style.background = 'var(--clr-brand, #B6FF8F)';
            btn.style.color = '#1a1a1a';
            btn.style.boxShadow = '0 2px 8px rgba(0,0,0,0.12)';
        });
    });

    // 2. Paraphrase Action [cite: 457]
    document.getElementById("paraphrase-btn").onclick = () => {
        const activeMode = document.querySelector('.mode-btn[data-active="true"]');
        const mode = activeMode ? activeMode.getAttribute('data-mode') : 'formal';
        sendPrompt(mode);
    };
    // document.getElementById("paraphrase-fast-btn").onclick = () => sendPrompt('fast');
    document.getElementById("test-btn").onclick = fetchFailedTest;
    // 3. Insert Action [cite: 470]
    document.getElementById("insert-btn").onclick = () => {
        // Logika insertText nanti disini
        //   ooxmlResult(),
        if (objects.math.length === 0 && objects.citations.length === 0) {
            insert(finalParaphrased)
        } else {
            insertxml(finalOoxml)
        }
        showNotification("Teks berhasil disisipkan!");
    };
    // 4. Retry Action [cite: 468]
    document.getElementById("copy-btn").onclick = async () => {
        try {
            await navigator.clipboard.writeText(paraphrasedRaw);
            showNotification("Teks berhasil disalin!");
        } catch (err) {
            console.error("Gagal menyalin teks: ", err);
        }
    };
}


// ================== NEW: LOGIKA HEALTH CHECK & FEEDBACK ==================

function updateHealthCheckUI(parapluie, bleu) {
    // Asumsi skor dari backend dikirim sebagai float 0.0 - 1.0
    // Konversi ke persen (0-100)
    const pScore = Math.round((parapluie || 0) * 100);
    const bScore = Math.round((bleu || 0) * 100);

    const pBar = document.getElementById('score-makna-bar');
    const pText = document.getElementById('score-makna-text');
    pBar.style.width = `${pScore}%`;

    // Logika Warna Akurasi Makna (ParaPLUIE)
    if (pScore >= 80) {
        pBar.className = "bg-green-500 h-2 rounded-full transition-all duration-500";
        pText.className = "font-bold text-green-600 dark:text-green-400";
        pText.innerText = `${pScore}% (Aman)`;
    } else if (pScore >= 60) {
        pBar.className = "bg-orange-500 h-2 rounded-full transition-all duration-500";
        pText.className = "font-bold text-orange-600 dark:text-orange-400";
        pText.innerText = `${pScore}% (Cek Ulang)`;
    } else {
        pBar.className = "bg-red-500 h-2 rounded-full transition-all duration-500";
        pText.className = "font-bold text-red-600 dark:text-red-400";
        pText.innerText = `${pScore}% (Berisiko)`;
    }

    const bBar = document.getElementById('score-variasi-bar');
    const bText = document.getElementById('score-variasi-text');
    bBar.style.width = `${bScore}%`;

    // Logika Warna Variasi Kata (BLEU) -> Asumsi semakin rendah metrik BLEU (sedikit plagiat), variasi semakin tinggi/bagus
    if (bScore <= 40) {
        bBar.className = "bg-green-500 h-2 rounded-full transition-all duration-500";
        bText.className = "font-bold text-green-600 dark:text-green-400";
        bText.innerText = `${100 - bScore}% (Bervariasi)`;
    } else {
        bBar.className = "bg-orange-500 h-2 rounded-full transition-all duration-500";
        bText.className = "font-bold text-orange-600 dark:text-orange-400";
        bText.innerText = `${100 - bScore}% (Mirip Asli)`;
    }
}

// Global state untuk 4 dimensi rating
let currentRating = {
    semantic: 0,
    gramatical: 0,
    syntactic: 0,
    lexical: 0,
};

// Konfigurasi 4 grup star rating sesuai HTML (id="star-rating-*")
const RATING_GROUPS = [
    { key: 'semantic',   containerId: 'star-rating-semantic'   },
    { key: 'gramatical', containerId: 'star-rating-gramatical' },
    { key: 'syntactic',  containerId: 'star-rating-syntactic'  },
    { key: 'lexical',    containerId: 'star-rating-lexical'    },
];

function setupFeedbackInteraction() {
    const submitBtn = document.getElementById('submit-feedback-btn');

    // Setup setiap grup bintang secara independen
    RATING_GROUPS.forEach(({ key, containerId }) => {
        const container = document.getElementById(containerId);
        if (!container) return;
        const btns = container.querySelectorAll('.star-btn');

        btns.forEach(btn => {
            btn.addEventListener('click', (e) => {
                const selected = parseInt(e.currentTarget.getAttribute('data-value'));
                currentRating[key] = selected;

                // Warnai bintang di grup ini saja
                btns.forEach(b => {
                    const val = parseInt(b.getAttribute('data-value'));
                    const icon = b.querySelector('svg');
                    if (icon) {
                        if (val <= selected) {
                            icon.classList.remove('text-gray-300');
                            icon.classList.add('text-yellow-400', 'fill-yellow-400');
                        } else {
                            icon.classList.add('text-gray-300');
                            icon.classList.remove('text-yellow-400', 'fill-yellow-400');
                        }
                    }
                });
            });
        });
    });

    // Interaksi klik Kirim
    if (submitBtn) {
        submitBtn.onclick = async () => {
            // Validasi: semua 4 dimensi harus diisi
            const missing = RATING_GROUPS.filter(({ key }) => currentRating[key] === 0);
            if (missing.length > 0) {
                showNotification("Silakan berikan semua penilaian bintang terlebih dahulu.");
                return;
            }

            if (!currentLogId) {
                showNotification("Belum ada hasil parafrase untuk dinilai.");
                return;
            }

            submitBtn.innerText = "Mengirim Penilaian...";
            submitBtn.disabled = true;
            submitBtn.classList.add('opacity-50');

            const comment = document.getElementById('feedback-comment').value;

            console.log("Mengirim feedback ke database...");
            console.log({ ...currentRating, comment, logId: currentLogId });

            try {
                const FEEDBACK_URL = `${BACKEND_URL}/${currentLogId}/feedback`;
                const response = await fetch(FEEDBACK_URL, {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        semantic:   currentRating.semantic,
                        gramatical: currentRating.gramatical,
                        syntactic:  currentRating.syntactic,
                        lexical:    currentRating.lexical,
                        userComment: comment || null,
                    })
                });

                if (!response.ok) {
                    const errData = await response.json().catch(() => ({}));
                    throw new Error(errData.error || `HTTP ${response.status}`);
                }

                console.log("Feedback berhasil disimpan!");

                // Ubah UI tombol setelah terkirim
                submitBtn.innerText = "\u2713 Penilaian Terkirim";
                submitBtn.classList.add('bg-green-100', 'text-green-700', 'dark:bg-green-900', 'dark:text-green-300');
                submitBtn.disabled = true;

                // Disable semua bintang di semua grup
                RATING_GROUPS.forEach(({ containerId }) => {
                    const container = document.getElementById(containerId);
                    if (!container) return;
                    container.querySelectorAll('.star-btn').forEach(btn => {
                        btn.disabled = true;
                        btn.classList.add('opacity-50');
                    });
                });

                document.getElementById('feedback-comment').disabled = true;
                showNotification("Penilaian berhasil disimpan!");
            } catch (err) {
                console.error("Gagal mengirim feedback:", err);
                submitBtn.innerText = "Kirim Penilaian";
                submitBtn.disabled = false;
                submitBtn.classList.remove('opacity-50');
                showNotification("Gagal mengirim penilaian: " + err.message);
            }

            // Reset state
            currentRating = { semantic: 0, gramatical: 0, syntactic: 0, lexical: 0 };
        };
    }
}

// Panggil setup interaksi saat file ter-load
document.addEventListener("DOMContentLoaded", () => {
    setupFeedbackInteraction();
});

// Simulasi Proses Parafrase (Mock Data)
async function runParaphraseSimulation(originalText, object) {
    // const inputSection = document.getElementById("input-section");
    const loadingState = document.getElementById("loading-state");
    const resultSection = document.getElementById("result-section");
    const btn = document.getElementById("paraphrase-btn");

    // UI Transition: Start Loading
    btn.disabled = true;
    btn.classList.add("opacity-50");
    loadingState.classList.remove("hidden");
    resultSection.classList.add("hidden");

    // Simulate API Delay (2 seconds)
    setTimeout(() => {
        loadingState.classList.add("hidden");
        resultSection.classList.remove("hidden");
        btn.disabled = false;
        btn.classList.remove("opacity-50");

        // Tampilkan Hasil Mockup Diff View 
        // Merah coret untuk kata lama, Hijau tebal untuk kata baru
        const mockResultHTML = `
          <span class="text-gray-400 line-through decoration-red-400 decoration-2">Menurut</span> 
          <span class="text-green-600 dark:text-green-400 font-bold bg-green-50 dark:bg-green-900/30 px-1 rounded">Berdasarkan pendapat</span> 
          Alvin, (2024), 
          <span class="text-gray-400 line-through decoration-red-400 decoration-2">rumus</span> 
          <span class="text-green-600 dark:text-green-400 font-bold bg-green-50 dark:bg-green-900/30 px-1 rounded">persamaan</span> 
          E=mc^2 
          <span class="text-gray-400 line-through decoration-red-400 decoration-2">adalah representasi</span> 
          <span class="text-green-600 dark:text-green-400 font-bold bg-green-50 dark:bg-green-900/30 px-1 rounded">merepresentasikan</span> 
          energi.
      `;

        document.getElementById("result-content").innerHTML = mockResultHTML;

        // Simulasi Warning jika skor rendah (acak untuk demo) 
        const warningBox = document.getElementById("warning-box");
        const randomScore = Math.random();
        if (randomScore < 0.3) {
            warningBox.classList.remove("hidden");
        } else {
            warningBox.classList.add("hidden");
        }

    }, 2000);
}

function showNotification(msg) {
    console.log(msg);
    const toast = document.getElementById('toast');
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}