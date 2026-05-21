# DoubleCMD MT Middleware

Style Guide'a göre EN→TR çeviri yapan servis. Phrase TMS ile entegre çalışır.

## Deploy (Railway - Ücretsiz)

1. [railway.app](https://railway.app) → "New Project" → "Deploy from GitHub repo"
2. Bu klasörü GitHub'a yükle
3. Railway'de Environment Variable ekle:
   - Key: `ANTHROPIC_API_KEY`
   - Value: `sk-ant-...` (Anthropic API key'in)
4. Deploy tamamlanınca bir URL alacaksın: `https://xxx.railway.app`

## Phrase'e Bağlama

1. Phrase → Settings → Integrations → Machine Translation Engines → Create
2. Engine type: **LibreTranslate**
3. API URL: `https://xxx.railway.app`
4. API Key: boş bırak (gerekmiyor)
5. Kaydet

## Test

```bash
curl -X POST https://xxx.railway.app/translate \
  -H "Content-Type: application/json" \
  -d '{"q": "Save file", "source": "en", "target": "tr"}'
```

Beklenen cevap:
```json
{"translatedText": "Dosyayı kaydet"}
```
