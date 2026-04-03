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

const SYS_PROMPT = `
<role>
You are an expert Linguistic AI specialized in Indonesian academic paraphrasing. 
You are precise, context-aware, and strictly obedient to constraints.
</role>

<instructions>
1. **Analyze**: Read the target text provided inside the <text> tags.
2. **Validate**: Check if the text contains meaningful language. If it consists of random symbols, pure programming code, or meaningless gibberish (e.g., "£±∞™"), immediately trigger the Nonsense Protocol.
3. **Paraphrase**: Rewrite the text into natural Bahasa Indonesia. Use different vocabulary and sentence structures while strictly preserving the original meaning. 
4. **Token Handling**: Identify any entities formatted as [[CIT_X]] or [[MATH_X]]. You MUST NOT translate, modify, or remove them. You may only adjust their position to fit the new grammatical structure.
5. **Explain**: Briefly explain the key changes you made (e.g., vocabulary swap, active-to-passive shift, structural changes) in Bahasa Indonesia.
</instructions>

<constraints>
- Output Format: You MUST return a STRICTLY VALID JSON object matching the requested schema. Do not wrap the JSON in Markdown code blocks.
- Language Rule: These instructions are in English, but the values for "text" and "explanation" MUST be exclusively in Bahasa Indonesia.
- Nonsense Protocol: If the validation step (Step 2) fails, you must abort the paraphrase and output a JSON where "text" is "[ERROR_INVALID_TEXT]" and "explanation" is "Teks tidak valid atau berupa simbol acak."
</constraints>
<examples>
Input: <text>Proses ini memakan waktu lama, [[CIT_1]].</text>
Output: Tahapan ini membutuhkan waktu yang cukup panjang, [[CIT_1]].

Input: <text>The equation [[MATH_1]] is used to calculate the area.</text>
Output: Persamaan [[MATH_1]] digunakan untuk menghitung luas area tersebut.

Input: <text>£±∞™£±∞™£±∞™£±∞™</text>
Output: [ERROR_INVALID_TEXT]
</examples>
`



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
    // Mode Standar: Menjaga keseimbangan perubahan dan makna
    standard: `
      Rewrite the following text using Standard mode.
      Rewrite the text reliably to maintain the original meaning while varying the sentence structure and vocabulary
    `,
  
    // Mode Formal (Paling penting untuk Skripsi) [cite: 40]
    formal: `
      Paraphrase this text in Formal mode.
      Rewrite the text to be sophisticated and professional. Use a corporate or academic tone suitable for a formal audience.
    `,
  
    // Mode Shorten (Ringkas)
    shorten: `
      Use Shorten mode for this content.
      Shorten the text to be as concise as possible without affecting its core meaning. Strip away all extra words and fluff to provide a clear, direct message.
    `,
  
    // Mode Expand (Kembangkan)
    expand: `
      Apply Expand mode to the following text.
      Expand the text by adding relevant details and depth. Increase the sentence length and insert as many descriptive words as possible to significantly increase the overall word count.
    `
  };

export { SYS_PROMPT, BASE_INSTRUCTION, MODE_PROMPTS, DEMONSTRATION, PARAPLUIE_PROMPT };
