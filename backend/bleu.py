from nltk.translate.bleu_score import sentence_bleu

def hitung_self_bleu(sumber, parafrase):
    # Tokenisasi (memecah kalimat jadi kata)
    reference = [sumber.lower().split()]
    hypothesis = parafrase.lower().split()

    # Menghitung skor BLEU
    return sentence_bleu(reference, hypothesis)

# Contoh penggunaan
sumber = """
Pemanfaatan teknologi Large Language Model (LLM) menawarkan potensi besar untuk membantu aspek parafrase dan penyempurnaan naskah [[CIT_1]]. Meskipun demikian, penerapan langsung LLM dalam penulisan karya ilmiah memiliki risiko teknis yang signifikan. Sifat dasar LLM yang bekerja secara probabilistik (memprediksi token kata) berpotensi menyebabkan "halusinasi" pada data yang seharusnya bersifat mutlak (deterministic), seperti mengubah tahun sitasi, memodifikasi variabel dalam rumus matematika, atau merusak format rich text (seperti cetak miring pada istilah asing) [[CIT_2]]. Selain itu, hasil parafrase generatif sering kali mengalami pergeseran makna (semantic drift) yang sulit dideteksi tanpa metode validasi kuantitatif [[CIT_3]].
"""
hasil = """
Penggunaan Large Language Model (LLM) memiliki potensi luas untuk membantu proses parafrase dan penyempurnaan dokumen [[CIT_1]], namun penerapannya secara langsung dalam penulisan ilmiah membawa risiko teknis yang cukup besar. Sifat probabilistik dari LLM dapat menyebabkan kesalahan pada data yang bersifat tetap, seperti tahun sitasi, variabel matematika, maupun format teks khusus [[CIT_2]]. Selain itu, hasil parafrase sering kali mengalami pergeseran makna yang sulit dideteksi tanpa adanya metode validasi yang terukur [[CIT_3]].
"""
score = hitung_self_bleu(sumber, hasil)
print(f"Self-BLEU: {(score)}")