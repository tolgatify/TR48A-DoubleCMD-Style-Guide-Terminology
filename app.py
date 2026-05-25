from flask import Flask, request, jsonify
import os
import anthropic

app = Flask(__name__)

# ── Style Guide ────────────────────────────────────────────────────────────────
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
- Use "kişi" or "kişiler" for neutral references.

=== GRAMMAR RULES ===

ABBREVIATIONS:
- Acronyms: uppercase, no periods between letters (TDK not T.D.K.), except T.C.
- Storage units always capitalized: KB, MB, GB, TB
- Non-breaking space between number and unit: 10 MB, %d kg
- Suffixes on acronyms use apostrophe based on pronunciation: TDK'nin, API'sı, SMTP'ye
- Do NOT abbreviate words unless strict UI character limit forces it.
- If abbreviation is necessary: Vowel Dropping (mesaj → msj) or Truncation with period (Komut → Kom.)
- NEVER abbreviate product names, titles, or headings.

CAPITALIZATION:
- Sentence-style: capitalize only first word and proper nouns.
- UI Buttons & Main Menus (1–3 words): Title Case (e.g., Ayarları Kaydet, Veriyi Dışa Aktar).
- Settings Headers & Dialog Titles: Title Case (e.g., Bildirim Yönetimi).
- Checkboxes, Radio Buttons & Option Descriptions: Sentence Case (e.g., Biri benden bahsettiğinde bana bildir.).
- File extensions: lowercase (.docx not .DOCX)

NUMBERS & FORMATS:
- Decimals: comma (1.250,50); thousands: period
- Date: DD.MM.YYYY (25.05.2026)
- Time: 24-hour (14:30)
- Percent: %100 (sign before number, no space)
- Currency: 1.250TL (no space between number and TL)

PUNCTUATION:
- Buttons/menu items: NO period at end (Kaydet, not Kaydet.)
- Full sentences in messages: always end with period.
- Ellipsis (...): only when action opens a new window/modal (Farklı Kaydet...)
- Exclamation (!): only for critical warnings, very sparingly.
- No comma before/after ve, veya, ya da.
- Do not use "ki" to connect clauses.
- Do not start standalone UI strings with conjunctions.

NOUNS:
- After a number/placeholder: always singular (3 dosya silindi, NOT 3 dosyalar silindi)
- %d items → %d dosya (singular)
- Do NOT translate brand/product names unless they have an established Turkish equivalent.
- Use apostrophe for suffixes on proper nouns: Windows'ta, DoubleCMD'de

VERBS:
- Buttons: imperative form (Gönder, Sil) — NEVER infinitive (Göndermek, Silmek)
- Progress/loading: use -yor suffix (Yükleniyor..., Eşitleniyor...)
- Error messages: use passive/neutral (Şifre hatalı. NOT Yanlış şifre girdiniz.)

GENITIVE:
- Menu items/UI labels: Undefined Noun Phrase (Dosya Özellikleri NOT Dosyanın Özellikleri)
- Apostrophe for proper nouns with suffixes: Windows'un (with suffix needs apostrophe), but Windows klasörü (no apostrophe without suffix)
- Avoid long nested genitive chains.

CONJUNCTIONS:
- Do not start standalone UI strings with conjunctions.
- "and/or" → ve/veya
- Avoid "ki" conjunction.

PLACEHOLDERS:
- Keep placeholder order (%s, %d, %1, %2) as in source unless Turkish grammar requires reordering.
- NEVER attach suffixes to placeholders: "%s silinemiyor" NOT "%s'i silinemiyor"
- Percent with placeholder: %%%d (sign before number)
- Placeholders with numbers: noun always singular (%d dosya, not %d dosyalar)

=== ERROR MESSAGES ===
- Sound natural, empathetic, human — not robotic.
- Prioritize natural Turkish flow over English word order.
- "Cannot/Could not": → yapılamıyor / yapılamadı (passive)
- "Failed to": → yapılamadı (passive)
- "Cannot find": → bulunamıyor / bulunamadı (passive)
- "Insufficient memory": → Bellek yetersiz / Bellek yeterli değil
- Use passive/neutral voice, never blame the user.
- Restructure sentence to follow logical chronological order of events.

=== KEY NAMES ===
Backspace→Geri Al, Tab→Sekme, Spacebar→Ara çubuğu, Up Arrow→Yukarı Ok,
Down Arrow→Aşağı Ok, Left Arrow→Sol Ok, Right Arrow→Sağ Ok,
Ctrl→Control, Windows key→Windows tuşu, Menu key→Menü tuşu
(All others stay in English: Alt, Enter, Esc, Delete, Insert, Home, End, Shift, Pause, etc.)

=== OUTPUT RULES ===
- Return ONLY the Turkish translation. No explanations, no notes, no alternatives.
- Do not add quotes around the translation.
- Preserve any placeholders (%s, %d, %1, %2, etc.) exactly as they appear.
- Preserve any formatting tags if present.
"""

# ── Terminology: GitHub raw dosyasından veya local Excel'den yükle ──────────
TERM_TEXT = ""

def load_terminology_from_github():
    """GitHub'daki terminoloji dosyasını yükle (ileride eklenebilir)."""
    # Repo'da henüz terminoloji CSV/TSV dosyası yok.
    # Dosya eklendiğinde buraya raw URL gelecek, örneğin:
    # https://raw.githubusercontent.com/tolgatify/TR48A-DoubleCMD-Style-Guide-Terminology/main/termbase.tsv
    pass

def load_terminology_from_excel(path="termbase mergan.xlsx"):
    """Local Excel termbase'den terminoloji yükle."""
    try:
        import openpyxl
        if not os.path.exists(path):
            print(f"Terminoloji dosyası bulunamadı: {path}")
            return ""
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        lines = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            en_term = str(row[0]).strip() if row[0] is not None else ""
            tr_term = str(row[1]).strip() if row[1] is not None else ""
            if en_term and tr_term and en_term != "None" and tr_term != "None":
                lines.append(f"{en_term} -> {tr_term}")
        if lines:
            print(f"BAŞARILI: {len(lines)} terim Excel'den okundu.")
            return (
                "\n\n=== DOUBLECMD TERMINOLOGY GLOSSARY ===\n"
                "Aşağıdaki terimleri her zaman sağ tarafında belirtilen Türkçe karşılıklarıyla çevir:\n\n"
                + "\n".join(lines)
            )
    except ImportError:
        print("openpyxl yüklü değil, terminoloji atlanıyor.")
    except Exception as e:
        print(f"Terminoloji Excel'den okunamadı: {e}")
    return ""

TERM_TEXT = load_terminology_from_excel()
FULL_INSTRUCTION = STYLE_GUIDE + TERM_TEXT

# ── Anthropic istemcisi (uygulama başlarken bir kez oluştur) ─────────────────
def get_client():
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY environment variable eksik")
    return anthropic.Anthropic(api_key=api_key)

CLIENT = None

def client():
    global CLIENT
    if CLIENT is None:
        CLIENT = get_client()
    return CLIENT

# ── Çeviri fonksiyonu ─────────────────────────────────────────────────────────
def translate_text(text: str) -> str:
    """Tek bir metni Claude Haiku ile çevir."""
    if not text or not text.strip():
        return ""
    message = client().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        system=FULL_INSTRUCTION,
        messages=[
            {
                "role": "user",
                "content": f"Translate the following English text to Turkish:\n\n{text}"
            }
        ]
    )
    return message.content[0].text.strip()

# ── Endpoints ─────────────────────────────────────────────────────────────────
@app.route("/translate", methods=["POST"])
def translate():
    data = request.get_json()
    if not data or "q" not in data:
        return jsonify({"error": "Missing 'q' field"}), 400

    source_text = data["q"]
    if not source_text or not source_text.strip():
        return jsonify({"translatedText": ""}), 200

    try:
        result = translate_text(source_text)
        return jsonify({"translatedText": result})
    except Exception as e:
        print(f"Translation Error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/phrase-mt", methods=["POST"])
def phrase_mt():
    data = request.get_json()
    texts = data.get("texts", [])
    if not texts:
        return jsonify({"translations": []}), 200

    try:
        translations = []
        for text in texts:
            translations.append(translate_text(text) if text.strip() else "")
        return jsonify({"translations": translations})
    except Exception as e:
        print(f"Phrase MT Error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/languages", methods=["GET"])
def languages():
    return jsonify([
        {"code": "en", "name": "English"},
        {"code": "tr", "name": "Turkish"}
    ])


@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "service": "DoubleCMD MT - TR48A",
        "model": "claude-haiku-4-5-20251001",
        "terminology_terms": len(TERM_TEXT.split("\n")) if TERM_TEXT else 0
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
