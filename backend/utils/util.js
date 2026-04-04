import fs from 'fs';
import path from 'path';
import { SYS_PROMPT, PARAPLUIE_PROMPT, MODE_PROMPTS } from "./prompt.js";
import { GoogleGenAI, ThinkingLevel } from '@google/genai';
import { bleu } from 'bleu-score';
import prisma from './prisma.js';

// --- THE VERCEL FIX ---
// If we are in Vercel (production) and have the JSON string in an environment variable
if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    // 1. Define a temporary file path
    const tempKeyPath = path.join('/tmp', 'google-credentials.json');
    
    // 2. Write the JSON string to this temporary physical file
    fs.writeFileSync(tempKeyPath, process.env.GOOGLE_APPLICATION_CREDENTIALS);
    
    // 3. Point the standard Google environment variable to this new temp file
    process.env.GOOGLE_APPLICATION_CREDENTIALS = tempKeyPath;
}
// ----------------------

// const GOOGLE_CLOUD_PROJECT = process.env.GOOGLE_CLOUD_PROJECT;
const GOOGLE_CLOUD_PROJECT = process.env.GOOGLE_CLOUD_PROJECT;
const GOOGLE_CLOUD_LOCATION = process.env.GOOGLE_CLOUD_LOCATION || 'global';

/**
 * Menghilangkan teks placeholder dari input teks.
 * @param {string} text - Teks yang akan dibersihkan
 * @returns {string} Teks tanpa placeholder
 */
export function removePlaceholderText(text) {
    if (!text || typeof text !== 'string') return text;
    const placeholders = [
        'Klik atau ketuk di sini untuk memasukkan teks.',
        'Click or tap here to enter text.'
    ];
    let result = text;
    for (const placeholder of placeholders) {
        result = result.split(placeholder).join('').trim();
    }
    return result;
}

/**
 * @param {string} sourceText        - Teks yang sudah ditokenisasi (input ke model)
 * @param {object} objects            - Mapping token → raw (citations, math, footnotes)
 * @param {string} paraphraseLogId    - ID ParaphraseLog yang harus sudah dibuat di server.js
 * @param {string} originalText       - Teks asli pengguna (disimpan ke kolom sourceText)
 * @param {string|null} tokenizedText - Teks bertokenisasi (disimpan ke kolom tokenizedInput)
 * @param {number} iterationNumber    - Nomor iterasi saat ini (1-based)
 */
export async function paraphrase(sourceText, objects, paraphraseLogId, originalText, tokenizedText, iterationNumber = 1, mode) {
    // console.log(`
    //     ---------------------MODE--------------------
    //     ${MODE_PROMPTS['formal']}
    // `)
    const userPrompt = `
<context>
${sourceText}
</context>

<task>
Generate a paraphrase of the <context> text using different words and sentence structures while still conveying the same meaning.

Strict Constraints:
1. Token Preservation: Ensure all [[CIT_X]] and [[MATH_X]] tokens remain exactly as they are.
2. Paragraph Structure: You MUST mirror the exact paragraph structure of the source text. 
   - If the source text contains multiple paragraphs separated by line breaks, your paraphrased output MUST maintain that exact same number of paragraphs, separated by "\\n".
   - STRICT RULE 1: Do NOT combine multiple paragraphs into a single block of text. 
   - STRICT RULE 2: Do NOT add "\n" at the end of individual sentences within a paragraph. Only use "\n" to separate actual paragraphs.
</task>

<mode>
${MODE_PROMPTS['formal']}
</mode>

<final_instruction>
Remember to think step-by-step about the sentence structure and vocabulary changes. 
However, your internal thoughts MUST NOT be printed outside the JSON. Summarize your reasoning ONLY inside the "explanation" field. 
</final_instruction>
`;
    const sysPrompt = SYS_PROMPT;

    const responseSchema = {
        type: "OBJECT",
        properties: {
            text: { type: "STRING" }
        },
        required: ["text"]
    };
    const client = new GoogleGenAI({
        vertexai: true,
        project: GOOGLE_CLOUD_PROJECT,
        location: GOOGLE_CLOUD_LOCATION,
        // googleApplicationCredentials: JSON.parse(process.env.GOOGLE_APPLICATION_CREDENTIALS),
    });

    const model = mode === 'fast' ? 'gemini-3.1-flash-lite-preview' : 'gemini-3-flash-preview';

    const paraphraseStart = Date.now();
    const response = await client.models.generateContent({
        model: model,
        config: {
            systemInstruction: sysPrompt,
            temperature: 1,
            responseMimeType: 'application/json',
            responseSchema: responseSchema,
        },
        contents: [{ role: 'user', parts: [{ text: userPrompt }] }]
    });
    const paraphraseLatency = Date.now() - paraphraseStart;

    // Extract paraphrase token usage
    const paraphraseUsage = response.usageMetadata || {};
    const paraphrase_promptTokens = paraphraseUsage.promptTokenCount ?? null;
    const paraphrase_completionTokens = paraphraseUsage.candidatesTokenCount ?? null;
    const paraphrase_totalTokens = paraphraseUsage.totalTokenCount ?? null;
    if (!response || !response.text) {
        throw new Error('Tidak ada respons dari model');
    }
    const jsonResponse = JSON.parse(response.text);

    let paraphrasedText = jsonResponse.text.trim();
    console.log(paraphrasedText)
    if (paraphrasedText == "[ERROR_INVALID_TEXT]") {
        return paraphrasedText
    }
    const paraphrased = getCleanText(paraphrasedText, objects)
    const paraPLUIEResult = await paraPLUIE(sourceText, paraphrased)
    const paraPLUIE_score = paraPLUIEResult.score;
    const BLEU_score = computeBLEU(sourceText, paraphrased)
    console.log("hasil parafrase: ", paraphrased)
    console.log("ParaPLUIE score: ", paraPLUIE_score)
    console.log("BLEU score: ", BLEU_score)

    // ── Simpan setiap hasil generasi ke tabel Paraphrase ──────────────────────
    if (paraphraseLogId) {
        try {
            await prisma.paraphrase.create({
                data: {
                    paraphraseLog_id:              paraphraseLogId,
                    sourceText:                    originalText ?? sourceText,
                    tokenizedInput:                tokenizedText ?? null,
                    paraphrasedText:               paraphrasedText,
                    iteration:                     iterationNumber,
                    bleuScore:                     BLEU_score ?? null,
                    paraPluieScore:                paraPLUIE_score ?? null,
                    paraphrase_promptTokens:       paraphrase_promptTokens ?? null,
                    paraphrase_completionTokens:   paraphrase_completionTokens ?? null,
                    paraphrase_totalTokens:        paraphrase_totalTokens ?? null,
                    paraphrase_latency:            paraphraseLatency ?? null,
                    parapluie_promptTokens:        paraPLUIEResult.promptTokens ?? null,
                    parapluie_completionTokens:    paraPLUIEResult.completionTokens ?? null,
                    parapluie_totalTokens:         paraPLUIEResult.totalTokens ?? null,
                    parapluie_latency:             paraPLUIEResult.latency ?? null,
                }
            });
            console.log(`✅ Paraphrase record saved (iteration ${iterationNumber})`);
        } catch (dbErr) {
            console.error('❌ Failed to save Paraphrase record:', dbErr);
        }
    }

    return {
        paraphrasedText: paraphrasedText,
        paraphrased: paraphrased,
        pScore: paraPLUIE_score,
        bScore: BLEU_score,
        paraphrase_promptTokens,
        paraphrase_completionTokens,
        paraphrase_totalTokens,
        paraphrase_latency: paraphraseLatency,
        parapluie_promptTokens: paraPLUIEResult.promptTokens,
        parapluie_completionTokens: paraPLUIEResult.completionTokens,
        parapluie_totalTokens: paraPLUIEResult.totalTokens,
        parapluie_latency: paraPLUIEResult.latency,
    }
}

export function getCleanText(tokenizedText, mappingObjects) {
    if (!tokenizedText || !mappingObjects) return tokenizedText;

    let cleanText = tokenizedText;

    if (mappingObjects.citations) {
        mappingObjects.citations.forEach(cit => {
            cleanText = cleanText.split(cit.token).join(cit.raw);
        });
    }

    if (mappingObjects.math) {
        mappingObjects.math.forEach(m => {
            cleanText = cleanText.split(m.token).join(m.raw);
        });
    }

    if (mappingObjects.footnotes) {
        mappingObjects.footnotes.forEach(f => {
            cleanText = cleanText.split(f.token).join(f.raw);
        });
    }

    return cleanText.trim();
}

async function paraPLUIE(S, H) {
    let probYes = -10000;
    let probNo = -10000;
    const client = new GoogleGenAI({
        vertexai: true,
        project: GOOGLE_CLOUD_PROJECT,
        location: GOOGLE_CLOUD_LOCATION,
    });

    const response_schema = { "type": "STRING", "enum": ["Yes", "No"] };
    const generationConfig = {
        maxOutputTokens: 1000,
        temperature: 0,
        responseMimeType: "text/x.enum",
        responseSchema: response_schema,
        responseLogprobs: true,
        logprobs: 5,
    };
    const prompt = PARAPLUIE_PROMPT + `
        (user): A: "${S}"; B: "${H}"
(assistant): `;

    try {
        const parapluieStart = Date.now();
        const response = await client.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: generationConfig
        });
        const parapluieLatency = Date.now() - parapluieStart;

        // Extract paraPLUIE token usage
        const parapluieUsage = response.usageMetadata || {};

        console.log(`📝 Jawaban Model: ${response.text}\n`);

        const candidate = response.candidates?.[0];
        const logprobsData = candidate?.logprobsResult?.topCandidates;

        if (logprobsData && logprobsData.length > 0) {
            console.log("📊 Analisis Probabilitas (Confidence Score):");
            console.log("-------------------------------------------");

            const firstTokenLogprobs = logprobsData[0].candidates;
            firstTokenLogprobs.forEach((tokenInfo) => {
                const tokenClean = tokenInfo.token.trim().toLowerCase();
                const prob = tokenInfo.logProbability;
                console.log("Token:", tokenInfo.token)
                console.log("logprob:", prob)
                if (tokenClean == 'yes' && prob > probYes) {
                    probYes = prob;
                } else if (tokenClean == 'no' && prob > probNo) {
                    probNo = prob;
                }
            });
        } else {
            console.log("⚠️ Logprobs tidak ditemukan di response.");
        }
        if (probYes != -10000 && probNo != -10000) {
            const paraPLUIEScore = probYes - probNo
            console.log("ParaPLUIE score: ", paraPLUIEScore)
            return {
                score: paraPLUIEScore,
                promptTokens: parapluieUsage.promptTokenCount ?? null,
                completionTokens: parapluieUsage.candidatesTokenCount ?? null,
                totalTokens: parapluieUsage.totalTokenCount ?? null,
                latency: parapluieLatency,
            }
        }
        // If logprobs couldn't determine score, return with metadata anyway
        return {
            score: undefined,
            promptTokens: parapluieUsage.promptTokenCount ?? null,
            completionTokens: parapluieUsage.candidatesTokenCount ?? null,
            totalTokens: parapluieUsage.totalTokenCount ?? null,
            latency: parapluieLatency,
        }
    } catch (error) {
        console.error("Error:", error);
        return {
            score: undefined,
            promptTokens: null,
            completionTokens: null,
            totalTokens: null,
            latency: null,
        }
    }
}

function computeBLEU(S, H) {
    const score = bleu(S, H, 4);
    return score;
}
