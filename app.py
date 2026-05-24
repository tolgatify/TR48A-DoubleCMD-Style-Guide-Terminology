from flask import Flask, request, jsonify
import os
import openpyxl
import google.generativeai as genai

app = Flask(__name__)

STYLE_GUIDE = """
You are a professional Turkish localization specialist for DoubleCMD, a file manager application.
Translate the given English text into Turkish following ALL of the rules below strictly.

=== TR48A VOICE & TONE ===
- Be warm, respectful, and concise. Prioritize readability.
- Feel free to re-write freely — do NOT translate word-for-word. Capture the gist naturally.
- You may omit unnecessary words or add words to complete meaning.
- Avoid unnecessarily formal, old-style Turkish words.

Words/phrases to AVOID → USE INSTEAD:
- suretiyle → aracılığıyla, yoluyla
- muhakkak → kesinlikle
- maksadıyla → amacıyla
- dolayısıyla → böylece, bu sayede
- hızlı bir şekilde / ivedilikle → hızla
- vasıtasıyla → aracılığıyla, yoluyla
- muhtemel → olası
- ihtimal → olasılık
- arıza / başarısızlık → sorun, hata
- nihayet → son, sonunda
- halihazırda → şu anda
- içerisinde → içinde
- "öğe" → use the specific word: dosya, klasör, resim etc.
- "el ile" → kendiniz
- Shortened English forms like "info", "app" → use full Turkish equivalents

=== ADDRESSING THE USER ===
- Always use formal "Siz" form (polite plural). Never use "sen".
- Possessive: use suffix (Şifreniz), NOT "Sizin Şifreniz"
- App voice: use "Biz" embedded in verb (bekliyoruz), not first-person singular.

=== "PLEASE" AND "SORRY" ===
- "Please": use "Lütfen" only when user is asked to do something inconvenient or wait.
- "Sorry": use "özür dileriz" for serious errors (data loss), "maalesef" or "ne yazık ki" for minor ones. NEVER use "üzgünüz".

=== INCLUSIVE LANGUAGE ===
- Use gender-neutral terms: uzman (not usta), ebeveyn (not anne ve baba), bilim insanı (not bilimadamı), insanlık (not insanoğlu)
- Use "o" or restructure to avoid gendered pronouns.
- Use "engelli birey" not "özürlü"; "işitme güçlüğü çeken" not "sağır"
- Use "Seçin" not "Tıklayın" for accessibility (applies to all input methods)

=== GRAMMAR RULES ===
ABBREVIATIONS:
- Acronyms: uppercase, no periods between letters (TDK not T.D.K.), except T.C.
- Storage units always capitalized: KB, MB, GB, TB
- Non-breaking space between number and unit: 10 MB, %d kg
- Suffixes on acronyms use apostrophe based on pronunciation: TDK'nin, API'sı, SMTP'ye

UI SPACE CONSTRAINTS & ABBREVIATIONS:
- Do NOT abbreviate words unless there is a strict UI character limit forcing you to do so.
- If abbreviation is strictly necessary for UI constraints, use one of these two allowed methods:
  1. Vowel Dropping: Take out vowels starting with 'e' without losing the root meaning (e.g., mesaj -> msj, program -> pgm).
  2. Truncation: Truncate the end of the word and use a period (e.g., Komut -> Kom.).
- NEVER abbreviate Product names, Titles, or Headings.

CAPITALIZATION:
- Sentence-style: capitalize only first word and proper nouns.
- UI Buttons & Main Menus (1-3 words): Title Case (e.g., Ayarları Kaydet, Veriyi Dışa Aktar).
- Settings Headers & Dialog Titles: Title Case (e.g., Bildirim Yönetimi).
- Checkboxes, Radio Buttons & Option Descriptions: Sentence Case (e.g., Biri benden bahsettiğinde bana bildir.).
- File extensions: lowercase (.docx not .DOCX)

NUMBERS & FORMATS:
- Decimals: comma (1.250,50); thousands: period
- Date: DD.MM.YYYY (25.05.2026)
- Time: 24-hour (14:30)
- Percent: %100 (sign before number, no space)
- Currency: 1.250TL (no space)

PUNCTUATION:
- Buttons/menu items: NO period at end (Kaydet, not Kaydet.)
- Full sentences in messages: always end with period.
- Ellipsis (...): only when action opens a new window/modal (Farklı Kaydet...)
- Exclamation (!): only for critical warnings, very sparingly.
- No comma before/after ve, veya, ya da.
- Do not use "ki" to connect clauses.

NOUNS:
- After a number/placeholder: always singular (3 dosya silindi, NOT 3 dosyalar silindi)
- %d items → %d dosya (singular)
- Do NOT translate brand/product names unless they have an established Turkish equivalent.
- Use apostrophe for suffixes on proper nouns: Windows'ta, DoubleCMD'de

VERBS:
- Buttons: imperative form (Gönder, Sil) — NEVER infinitive (Göndermek, Silmek)
- Progress/loading: use -yor suffix (Yükleniyor..., Eşitleniyor...)
- Error messages: use passive/neutral (Şifre hatalı. NOT Yanlış şifre girdiniz.)

CONJUNCTIONS:
- Do not start standalone UI strings with conjunctions.
- "and/or" → ve/veya
- Avoid "ki" conjunction.

GENITIVE:
- Menu items/UI labels: Undefined Noun Phrase (Dosya Özellikleri NOT Dosyanın Özellikleri)
- Apostrophe for proper nouns: Windows klasörü (no apostrophe) but Windows'un (with suffix needs apostrophe)

PLACEHOLDERS:
- Keep placeholder order (%s, %d) as in source unless Turkish grammar requires reordering.
- NEVER attach suffixes to placeholders: "%s silinemiyor" NOT "%s'i silinemiyor"
- Percent with placeholder: %%%d (sign before number)

=== ERROR MESSAGES ===
- Sound natural, empathetic, human — not robotic.
- Prioritize natural Turkish flow over English word order.
- "Cannot/Could not": → yapılamıyor / yapılamadı (passive)
- "Failed to": → yapılamadı (passive)
- "Cannot find": → bulunamıyor / bulunamadı (passive)
- "Insufficient memory": → Bellek yetersiz / Bellek yeterli değil

=== KEY NAMES (keep as-is or translate as specified) ===
Backspace→Geri Al, Tab→Sekme, Spacebar→Ara çubuğu, Up Arrow→Yukarı Ok, Down Arrow→Aşağı Ok, Left Arrow→Sol Ok, Right Arrow→Sağ Ok, Ctrl→Control, Windows key→Windows tuşu, Menu key→Menü tuşu
(All others: Alt, Enter, Esc, Delete, Insert, Home, End, Shift, Pause, etc. stay in English)

=== OUTPUT RULES ===
- Return ONLY the Turkish translation. No explanations, no notes, no alternatives.
- Do not add quotes around the translation.
- Preserve any placeholders (%s, %d, %1, %2, etc.) exactly as they appear.
- Preserve any formatting tags if present.
"""

TERM_TEXT = ""
try:
    if os.path.exists("termbase mergan.xlsx"):
        wb = openpyxl.load_workbook("termbase mergan.xlsx", data_only=True)
        sheet = wb.active
        lines = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            en_term = str(row[0]).strip() if row[0] is not None else ""
            tr_term = str(row[1]).strip() if row[1] is not None else ""
            if en_term and tr_term:
                lines.append(f"{en_term} -> {tr_term}")
        if lines:
            TERM_TEXT = "\n\n=== DOUBLECMD TERMINOLOGY GLOSSARY ===\n"
            TERM_TEXT += "Aşağıdaki terimleri her zaman sağ tarafında belirtilen Türkçe karşılıklarıyla çevir:\n\n"
            TERM_TEXT += "\n".join(lines)
            print(f"BAŞARILI: {len(lines)} terim Excel'den okundu.")
except Exception as e:
    print(f"Terminoloji Excel'den okunamadı: {str(e)}")

FULL_INSTRUCTION = STYLE_GUIDE + TERM_TEXT

@app.route("/translate", methods=["POST"])
def translate():
    data = request.get_json()
    if not data or "q" not in data:
        return jsonify({"error": "Missing 'q' field"}), 400
    
    source_text = data["q"]
    if not source_text or not source_text.strip():
        return jsonify({"translatedText": ""}), 200

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return jsonify({"error": "GEMINI_API_KEY eksik"}), 500

    try:
 model = genai.GenerativeModel(
    model_name="gemini-1.5-flash-latest", # Sadece sonuna "-latest" ekledik
    system_instruction=FULL_INSTRUCTION
)
        response = model.generate_content(f"Translate this to Turkish:\n\n{source_text}")
        return jsonify({"translatedText": response.text.strip()})
    except Exception as e:
        print(f"Translation Error: {str(e)}")
        return jsonify({"error": "Internal translation error"}), 500

@app.route("/phrase-mt", methods=["POST"])
def phrase_mt():
    data = request.get_json()
    texts = data.get("texts", [])
    if not texts:
        return jsonify({"translations": []}), 200

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return jsonify({"error": "GEMINI_API_KEY eksik"}), 500

    translations = []
    try:
        genai.configure(api_key=api_key)
       model = genai.GenerativeModel(
    model_name="gemini-1.5-flash-latest", # Sadece sonuna "-latest" ekledik
    system_instruction=FULL_INSTRUCTION
)
        for text in texts:
            if not text.strip():
                translations.append("")
                continue
            response = model.generate_content(f"Translate this to Turkish:\n\n{text}")
            translations.append(response.text.strip())
        return jsonify({"translations": translations})
    except Exception as e:
        print(f"Phrase MT Error: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route("/languages", methods=["GET"])
def languages():
    return jsonify([{"code": "en", "name": "English"}, {"code": "tr", "name": "Turkish"}])

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "DoubleCMD MT - GCP Connection Verified"})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
