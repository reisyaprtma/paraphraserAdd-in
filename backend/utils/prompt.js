import prisma from './prisma.js';

const BASE_INSTRUCTION = `
  Look at the samples of a sentence and its intelligible paraphrase:
  1. I don’t know if you are familiar with that. =>
  I have no idea if you’re familiar with that.
  2. what other long-range goals do you have besides college? =>
  Apart from college, what are your other long-term objectives?
  3. I don’t have access either. Although, I did at one time =>
  In the past, I had access, but currently, I don’t.
  4. Right now I’ve got it narrowed down to the top four teams. =>
  At this point, I’ve trimmed my options and picked 4 top teams.
  5. prohibition didn’t stop it and didn’t do anything really. =>
  It continued despite the prohibition, which didn’t accomplish anything.
`;

/**
 * Fetches top-scoring paraphrases from the DB to use as dynamic few-shot examples.
 * Filters by paraphrase mode via the related ParaphraseLog table.
 * Returns up to 5 unique (tokenizedInput → paraphrasedText) pairs with the best paraPluieScore.
 */
async function fetchTopParaphrases(mode) {
    try {
        const results = await prisma.$queryRaw`
            SELECT p."tokenizedInput", p."paraphrasedText"
            FROM "Paraphrase" p
            INNER JOIN "ParaphraseLog" pl ON p."paraphraseLog_id" = pl."id"
            WHERE p."paraPluieScore" > 0
              AND p."bleuScore" <= 0.7
              AND p."tokenizedInput" IS NOT NULL
              AND LENGTH(p."tokenizedInput") <= 400
              AND LENGTH(p."paraphrasedText") <= 400
              AND pl."mode" = ${mode}
            ORDER BY p."paraPluieScore" DESC
            LIMIT 5 
        `;
        return results;
    } catch (err) {
        console.error('⚠️ Failed to fetch top paraphrases for prompt:', err.message);
        return [];
    }
}

/**
 * Builds the system prompt dynamically, injecting real high-quality paraphrase
 * examples from the database alongside the static examples.
 */
async function buildSysPrompt(mode) {
    const topParaphrases = await fetchTopParaphrases(mode);
    console.log("topParaphrases: ", topParaphrases)

    // Build dynamic examples block from DB results
    let dynamicExamples = '';
    if (topParaphrases.length > 0) {
        dynamicExamples = '\n--- Real high-quality examples from past paraphrases ---\n';
        topParaphrases.forEach((row, i) => {
            dynamicExamples += `\nInput: <text>${row.tokenizedInput}</text>\n`;
            dynamicExamples += `Output: ${row.paraphrasedText}\n`;
        });
    }

    return `
<role>
You are an expert Linguistic AI specialized in Indonesian academic paraphrasing. 
You are precise, context-aware, and strictly obedient to constraints.
</role>

<instructions>
1. **Analyze**: Read the target text provided inside the <text> tags.
2. **Validate**: Check if the text contains meaningful language. If it consists of random symbols, pure programming code, or meaningless gibberish (e.g., "£±∞™"), immediately trigger the Nonsense Protocol.
3. **Paraphrase**: Rewrite the text into natural Bahasa Indonesia. Use different vocabulary and sentence structures while strictly preserving the original meaning. 
4. **Token Handling**: Identify any entities formatted as [[CIT_X]] or [[MATH_X]]. You MUST NOT translate, modify, or remove them. You may only adjust their position to fit the new grammatical structure.
</instructions>

<constraints>
- Output Format: You MUST return a STRICTLY VALID JSON object matching the requested schema. Do not wrap the JSON in Markdown code blocks.
- Language Rule: These instructions are in English, but the values for "text" MUST be exclusively in Bahasa Indonesia.
- Nonsense Protocol: If the validation step (Step 2) fails, you must abort the paraphrase and output a JSON where "text" is "[ERROR_INVALID_TEXT]"
</constraints>
<examples>
Input: <text>Proses ini memakan waktu lama, [[CIT_1]].</text>
Output: Tahapan ini membutuhkan waktu yang cukup panjang, [[CIT_1]].

Input: <text>The equation [[MATH_1]] is used to calculate the area.</text>
Output: Persamaan [[MATH_1]] digunakan untuk menghitung luas area tersebut.

Input: <text>£±∞™£±∞™£±∞™£±∞™</text>
Output: [ERROR_INVALID_TEXT]
${dynamicExamples}</examples>
`;
}



const DEMONSTRATION = `
  I don’t know if you are familiar with that, in other words I have no idea if you’re familiar with that.
  What other long-range goals do you have besides college?, in other words Apart from college, what are your other long-term objectives?
  I don’t have access either. Although, I did at one time, in other words In the past, I had access, but currently, I don’t.
  Right now I’ve got it narrowed down to the top four teams, in other words At this point, I’ve trimmed my options and picked 4 top teams.
  Prohibition didn’t stop it and didn’t do anything really, in other words It continued despite the prohibition, which didn’t accomplish anything.
`

const PARAPLUIE_PROMPT = `
(user): You will receive two sentences A and B. Do
these two sentences mean the same thing? Answer
with only one word "Yes" or "No".
(assistant): Please provide the sentences for me
to evaluate.
(user): A: "Amrozi accused his brother, whom he
called "the witness", of deliberately distorting his
evidence ."; B: "Amrozi accused his brother, whom
he disparagingly referred to as ’the liar witness’,
of intentionally twisting his testimony."
(assistant): No
(user): A: "Pennmakkal is an Indian Malayalam
film from 1966, produced by J. Sasikumar and
directed by KP Kottarakkara."; B: "The Indian
Malayalam film ’Pennmakkal’, released in 1966,
was produced by J. Sasikumar and directed by KP
Kottarakkara."
(assistant): Yes
(user): A: "Sorkin , who faces charges of conspiracy to obstruct justice and lying to a grand jury ,
was to have been tried separately."; B: "Despite
being accused of conspiring to obstruct justice and
perjury, Sorkin was supposed to stand trial on his
own."
(assistant): No
(user): A: "Gilroy police and FBI agents described
Gehring as cooperative , but said Saturday that he
had revealed nothing about what had happened
to the children ."; B: "Although Gilroy police and
FBI agents reported that Gehring was cooperative
, he hadn’t disclosed any information about the
children’s whereabouts or what had happened to
them as of Saturday ."
(assistant): No
(user): A: "Whereas “e” the electric charge of
the particle and A is the magnetic vector potential
of the electromagnetic field."; B: "The electric
charge of the particle is denoted by “e”, and the
magnetic vector potential of the electromagnetic
field is denoted by ’A’."
(assistant): Yes
(user): A: "The Jidanul River is a tributary of the
Jiul de Vest River in Romania."; B: "The Jidanul
River is a mere insignificant stream that flows into
the grand Jiul de Vest River in Romania."
(assistant): No
`

const MODE_PROMPTS = {
    // Mode Formal: Cocok untuk konteks akademis dan profesional
    formal: `
      Paraphrase this text in Formal mode.
      Rewrite the text using a formal, professional, and academic tone. The result must be suitable for academic or professional documents such as reports, theses, or official correspondence.
      Maintain the original meaning precisely. Avoid casual language, slang, or colloquial expressions.
    `,

    // Mode Academic: Fokus pada detail dan penggunaan kata yang tepat untuk konteks akademik
    academic: `
      Paraphrase this text in Academic mode.
      Rewrite the text with precise, discipline-appropriate vocabulary and a rigorous academic writing style. Pay careful attention to technical accuracy and the use of domain-specific terminology.
      Ensure the paraphrased result is suitable for scholarly publications, academic papers, or research writing. Preserve all factual details and the original argument structure.
    `
  };

export { buildSysPrompt, BASE_INSTRUCTION, MODE_PROMPTS, DEMONSTRATION, PARAPLUIE_PROMPT };
