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
const BACKEND_URL = "https://backend-server-paraphraser.vercel.app/api/paraphrase";
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
    if (!box) return;
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();

            const text = selection.text || '';

            if (text.length > 2500) {
                btn.disabled = true;
                btn.classList.add("opacity-50");
            }
            else if (text.length === 0) {
                btn.disabled = true;
                btn.classList.add("opacity-50");
            }
            else {
                btn.disabled = false;
                btn.classList.remove("opacity-50");
            }
            box.value = text;
            if (statusEl) {
                if (text.length > 2500) {
                    statusEl.innerHTML = `<i data-lucide="alert-triangle" class="text-red-500 w-3 h-3"></i> Teks melebihi batas panjang (${text.length} karakter)`;
                    // statusEl.classList.add('active');
                    statusEl.classList.remove('active');
                    statusEl.classList.add('bg-red-100');
                    statusEl.classList.add('text-red-500');
                } else if (text.length > 0) {
                    statusEl.innerHTML = `<i data-lucide="check" class="w-3 h-3"></i> ${text.length} karakter terseleksi`;
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

    await Word.run(async (context) => {
        // 1️⃣ SETUP & LOAD AWAL
        const selection = context.document.getSelection();
        const contentControls = selection.contentControls;
        selection.load("text");
        const html = selection.getHtml()

        // Load text agar bisa disimpan sebagai raw text sitasi
        contentControls.load("items/text, items/tag");
        const mainOoxmlResult = selection.getOoxml();
        await context.sync();
        const plainTextWithStructure = htmlToText(html.value, {
            wordwrap: false,
            selectors: [
                // { selector: 'h1', options: { uppercase: true } },
                { selector: 'table', options: { uppercaseHeader: true } },
                { selector: 'a', options: { ignoreHref: true } } // Biar nggak muncul link (http://...)
            ]
        });
        seleksi_teks = selection.text
        console.log(seleksi_teks)

        // 2️⃣ QUEUE OOXML REQUESTS
        const citationRequests = contentControls.items.map((cc) => {
            return {
                text: cc.text, // Simpan teks asli sitasi (misal: "(Anggita, 2024)")
                tag: cc.tag,
                ooxmlProxy: cc.getOoxml()
            };
        });

        await context.sync();

        // ---------------------------------------------------------
        // MULAI PEMROSESAN DATA
        // ---------------------------------------------------------

        let currentXml = mainOoxmlResult.value || "";
        if (!currentXml) return;

        // 3️⃣ PASS 1: MATH TOKENIZATION
        let mathIndex = 0;
        const oMathRegex = /<m:oMath[\s\S]*?<\/m:oMath>|<m:oMathPara[\s\S]*?<\/m:oMathPara>/gi;

        currentXml = currentXml.replace(oMathRegex, (match) => {
            const token = ` [[MATH_${mathIndex + 1}]] `;

            // 🔥 RAW TEXT MATH: Hapus semua tag XML agar sisa angka/hurufnya saja
            // Contoh: <m:t>E</m:t><m:t>=</m:t>... menjadi "E=mc^2"
            let rawMath = match.replace(/<[^>]+>/g, "").trim();

            mappingData.math.push({
                token: token.trim(),
                ooxml: match,
                raw: rawMath // <--- Simpan Raw Text Math
            });
            mathIndex++;
            return token;
        });

        // 4️⃣ PASS 2: CLEANING XML TO TEXT (Sama seperti sebelumnya)
        let textBase = currentXml
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

        // 5️⃣ PASS 3: CITATION TOKENIZATION
        citationRequests.forEach((req, index) => {
            const citToken = `[[CIT_${index + 1}]]`;

            mappingData.citations.push({
                token: citToken,
                ooxml: req.ooxmlProxy.value,
                raw: req.text // <--- Simpan Raw Text Sitasi dari Word
            });

            // Replace Text asli dengan Token
            textBase = textBase.split(req.text).join(citToken);
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
async function callAPI(originalText, tokenizedText, objects) {
    try {
        const payload = {
            originalText: originalText,
            tokenizedText: tokenizedText,
            objects: objects,

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

async function sendPrompt() {
    const diffToggle = document.getElementById("diff-toggle");
    const toggleCircle2 = document.getElementById("toggle-circle-2");
    const failedresultSection = document.getElementById("failed-results-container");
    diffToggle.classList.remove("active");
    toggleCircle2.classList.remove("translate-x-5")
    failedresultSection.classList.add("hidden")

    const starBtns = document.querySelectorAll('.star-btn');
    const submitBtn = document.getElementById('submit-feedback-btn');
    const comment = document.getElementById('feedback-comment');
    comment.value = "";
    comment.disabled = false;
    currentRating = 0; // Reset rating state
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
    // submitBtn.classList.add('bg-gray-100', 'hover:bg-gray-200');
    submitBtn.classList.remove('bg-green-100', 'text-green-700', 'dark:bg-green-900', 'dark:text-green-300');

    const warningBox = document.getElementById("warning-box")
    warningBox.classList.add('hidden')
    const errorBox = document.getElementById("error-box")
    errorBox.classList.add('hidden')

    const extractedToken = await extractCombinedTokens()
    tokenizedTexts = extractedToken.text
    objects = extractedToken.object
    console.log(tokenizedTexts)
    console.log(objects)
    const tokenizedText = tokenizedTexts
    try {
        // const inputSection = document.getElementById("input-section");
        const loadingState = document.getElementById("loading-state");
        const resultSection = document.getElementById("result-section");
        const btn = document.getElementById("paraphrase-btn");
        // const selectedMode = document.getElementById("paraphrase-mode")
        // UI Transition: Start Loading
        btn.disabled = true;
        btn.classList.add("opacity-50");
        loadingState.classList.remove("hidden");
        resultSection.classList.add("hidden");
        console.log(seleksi_teks)
        await callAPI(seleksi_teks, tokenizedText, objects)

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
                        // KASUS CITATION: Extract Body dulu
                        contentToInsert = extractBodyFromOoxml(mappingFound.ooxml);

                        // Buang kulit paragraf (<w:p>)
                        contentToInsert = contentToInsert
                            .replace(/<w:p [^>]*>/g, "")
                            .replace(/<w:p>/g, "")
                            .replace(/<\/w:p>/g, "")
                            .replace(/<w:sectPr[^>]*>[\s\S]*?<\/w:sectPr>/g, "");
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

    // Output ke DOM
    document.getElementById("result-content").innerHTML = unifiedHtml;
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
    if (document.getElementById("removal-count")) document.getElementById("removal-count").innerText = remCount;
    if (document.getElementById("addition-count")) document.getElementById("addition-count").innerText = addCount;
}

async function showResult(tokenizedResult, data, parapluie, bleu) {
    const loadingState = document.getElementById("loading-state");
    const resultSection = document.getElementById("result-section");
    const btn = document.getElementById("paraphrase-btn");

    loadingState.classList.add("hidden");
    resultSection.classList.remove("hidden");
    btn.disabled = false;
    btn.classList.remove("opacity-50");

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
    let mappingData = {
        math: [],
        citations: []
    };

    // 1. Reset & Show UI
    loadingState.classList.add("hidden");
    failedresultSection.classList.remove("hidden");
    btn.disabled = false;
    btn.classList.remove("opacity-50");

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
                unifiedHtml += `<span class="removal-highlight hidden" style="background-color:rgba(255, 8, 8, 0.05); color:rgb(96, 94, 94); text-decoration: line-through; text-decoration-color:red">${bufferRem}</span>`;
            }
            if (bufferAdd.trim().length > 0) {
                unifiedHtml += `<span class="addition-highlight" style="background-color: #ccffcc; color: #006600;">${bufferAdd}</span>`;
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

        const card = document.createElement("div");
        card.className = "card animate-slide-up";

        // LOGIKA HEALTH CHECK
        const pScore = pScores[index] || 0;
        const bScore = bScores[index] || 0;
        const similarityPercent = Math.round(bScore * 100);

        // Makna: < 0 berarti bergeser (Merah), >= 0 berarti aman (Hijau)
        const maknaStatus = pScore >= 0
            ? `<span class="text-green-600 dark:text-green-400 font-bold flex items-center gap-1"><i data-lucide="check-circle" class="w-4 h-4"></i> Makna Aman</span>`
            : `<span class="text-red-600 dark:text-red-400 font-bold flex items-center gap-1"><i data-lucide="alert-circle" class="w-4 h-4"></i> Makna Tidak Aman</span>`;

        // Variasi Kata (BLEU): Asumsi > 40% terlalu mirip (Merah), <= 40% aman (Hijau)
        // Lu bisa atur angka 40 ini sesuai threshold dosen lu nanti.
        const isMirip = similarityPercent > 40;
        const bleuStatus = isMirip
            ? `<div class="flex justify-between text-xs mb-1">
                   <span class="font-medium text-gray-700 dark:text-gray-300">Variasi Kata <span class="text-red-500 font-bold">(Kemiripan Tinggi)</span></span>
                   <span class="font-bold text-red-500">${similarityPercent}% Mirip</span>
               </div>
               <div class="w-full bg-gray-200 dark:bg-slate-700 rounded-full h-1.5">
                   <div class="bg-red-500 h-1.5 rounded-full" style="width: ${similarityPercent}%"></div>
               </div>`
            : `<div class="flex justify-between text-xs mb-1">
                   <span class="font-medium text-gray-700 dark:text-gray-300">Variasi Kata <span class="text-green-500">(Keragaman Baik)</span></span>
                   <span class="font-bold text-green-500">${similarityPercent}% Mirip</span>
               </div>
               <div class="w-full bg-gray-200 dark:bg-slate-700 rounded-full h-1.5">
                   <div class="bg-green-500 h-1.5 rounded-full" style="width: ${similarityPercent}%"></div>
               </div>`;

        card.innerHTML = `
            <div class="card-header">
                <span class="opsi-badge">Opsi ${index + 1}</span>
            </div>
            
            <div class="card-body" style="border-bottom:1px solid var(--clr-border)">
                <div style="font-size:11px;margin-bottom:6px">${maknaStatus}</div>
                <div>${bleuStatus}</div>
            </div>

            <div class="card-body">
                <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
                    <span class="section-label" style="margin-bottom:0">Diff View</span>
                    <button id="diff-toggle-failed-${index}" class="diff-toggle-btn">
                        <div id="toggle-circle-failed-${index}" class="diff-toggle-thumb"></div>
                    </button>
                </div>
                <div id="result-content">
                    ${unifiedHtml}
                </div>
            </div>
            
            <div class="card-footer">
                <div class="diff-count">
                    <span class="removed">− <span id="removal-count-failed-${index}">${remCount}</span></span>
                    <span class="added">+ <span id="addition-count-failed-${index}">${addCount}</span></span>
                </div>
                
                <div style="display:flex;gap:8px">
                    <button id="copy-btn-failed-${index}" class="btn-icon" title="Salin Teks">
                        <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                    </button>
                    <button id="insert-btn-failed-${index}" class="btn-icon filled" title="Sisipkan">
                        <i data-lucide="download" class="w-3.5 h-3.5"></i>
                    </button>
                </div>
            </div>
        `;

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
        if (insertBtn && ooxmlData && ooxmlData[index]) {
            insertBtn.addEventListener("click", function () {
                insertxml(ooxmlData[index]);
                showNotification("Teks opsi " + (index + 1) + " berhasil disisipkan.");
            });
        }
        // Tombol copy khusus untuk opsi ini
        const copyBtn = card.querySelector(`#copy-btn-failed-${index}`);
        if (copyBtn) {
            copyBtn.addEventListener("click", async function () {
                try {
                    await navigator.clipboard.writeText(paraphrasedArray[index]);
                    showNotification("Teks opsi " + (index + 1) + " berhasil disalin.");
                    console.log("Teks opsi " + (index + 1) + " disalin.");
                } catch (err) {
                    console.error("Gagal menyalin", err);
                }
            });
        }

        container.appendChild(card);
    });
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
    loadingState.classList.add("hidden");
    btn.disabled = false;
    btn.classList.remove("opacity-50");
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

    // Mock teks asli dan objects untuk rendering diff
    seleksi_teks = "Penggunaan teknologi kecerdasan buatan dalam penulisan karya ilmiah dapat membantu meningkatkan efisiensi waktu pengerjaan tugas akhir bagi mahasiswa tingkat akhir.";
    objects = { math: [], citations: [] };

    // UI: Show loading
    btn.disabled = true;
    btn.classList.add("opacity-50");
    loadingState.classList.remove("hidden");
    failedResultSection.classList.add("hidden");

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
    const darkModeToggle = document.getElementById("dark-mode-toggle");
    const warningBox = document.getElementById("warning-box")
    warningBox.classList.add('hidden')
    const errorBox = document.getElementById("error-box")
    errorBox.classList.add('hidden')
    const htmlEl = document.documentElement;
    const toggleCircle = document.getElementById("toggle-circle");

    settingsBtn.onclick = () => {
        settingsMenu.classList.toggle("hidden");
    };

    darkModeToggle.onclick = () => {
        htmlEl.classList.toggle("dark");
        if (htmlEl.classList.contains("dark")) {
            toggleCircle.classList.add("translate-x-5");
        } else {
            toggleCircle.classList.remove("translate-x-5");
        }
    };

    // 2. Paraphrase Action [cite: 457]
    document.getElementById("paraphrase-btn").onclick = sendPrompt;
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

// Global state untuk rating
let currentRating = 0;

function setupFeedbackInteraction() {
    const starBtns = document.querySelectorAll('.star-btn');
    const submitBtn = document.getElementById('submit-feedback-btn');
    // starBtns.forEach(btn => {
    //     btn.disabled = false;
    //     btn.classList.remove('opacity-50');
    // })
    // submitBtn.disabled = false;
    // submitBtn.classList.remove('opacity-50');

    starBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            currentRating = parseInt(e.currentTarget.getAttribute('data-value'));

            // Warnai SVG di dalam button
            starBtns.forEach(b => {
                const val = parseInt(b.getAttribute('data-value'));
                const icon = b.querySelector('svg'); // Target SVG hasil render Lucide

                if (icon) {
                    if (val <= currentRating) {
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

    // Interaksi klik Kirim
    if (submitBtn) {

        submitBtn.onclick = async () => {
            submitBtn.innerText = "Mengirim Penilaian...";
            submitBtn.disabled = true;
            submitBtn.classList.add('opacity-50');
            const comment = document.getElementById('feedback-comment').value;
            if (currentRating === 0) {
                // Bisa pakai alert atau showNotification buatanmu
                showNotification("Silakan berikan rating bintang terlebih dahulu.");
                return;
            }

            if (!currentLogId) {
                showNotification("Belum ada hasil parafrase untuk dinilai.");
                return;
            }

            console.log("Mengirim feedback ke database...");
            console.log({ rating: currentRating, comment: comment, logId: currentLogId });

            try {
                const FEEDBACK_URL = `${BACKEND_URL}/${currentLogId}/feedback`;
                const response = await fetch(FEEDBACK_URL, {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        rating: currentRating,
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
                submitBtn.classList.remove('bg-gray-100', 'hover:bg-gray-200');
                submitBtn.classList.add('bg-green-100', 'text-green-700', 'dark:bg-green-900', 'dark:text-green-300');
                submitBtn.disabled = true;
                starBtns.forEach(btn => {
                    btn.disabled = true;
                    btn.classList.add('opacity-50');
                });
                document.getElementById('feedback-comment').disabled = true;
                showNotification("Penilaian berhasil disimpan!");
            } catch (err) {
                console.error("Gagal mengirim feedback:", err);
                showNotification("Gagal mengirim penilaian: " + err.message);
            }

            // Reset state
            currentRating = 0;
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