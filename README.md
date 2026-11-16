# Logistika Telegram bot (BEPUL tarif)

Fayllar: `bot.py` (asosiy), `planner.py` (yuk rejasi), `database.py` (baza), `import_excel.py`, `migrate_sqlite_to_pg.py` (bir martalik ko'chirish).

## Qadamlar (oddiy tilda)
1. **Bot yarating**: Telegram'da @BotFather → /newbot → TOKEN oling. /mybots → Bot Settings → Group Privacy → **Turn off**.
2. **GitHub**: yangi repo yarating (**Public** — bepul ping uchun shart), shu papkani yuklang. `.gitignore` maxfiy fayllarni o'zi chiqarib tashlaydi. Token/parol hech qachon GitHub'ga yozilmaydi.
3. **Supabase** (baza): supabase.com → New project (parol oddiy harf-raqamdan bo'lsin). Connect → **Session pooler** manzilini nusxalang → bu `DATABASE_URL`.
4. **Kompyuteringizda** `.env` yarating (`.env.example` dan nusxa), qiymatlarni yozing. Keyin bir marta:
   `pip install -r requirements.txt` va `python migrate_sqlite_to_pg.py` — 54 ta yuk Supabase'ga o'tadi (qayta ishlatsangiz ham takrorlanmaydi).
5. **Render**: New → Web Service → GitHub repo → Free. Build: `pip install -r requirements.txt`, Start: `python bot.py`, Health Check Path: `/health`. Environment'ga: BOT_TOKEN, ADMIN_IDS, ANTHROPIC_API_KEY, DATABASE_URL, COLLECT_CHAT_ID, LOAD_CHAT_ID.
6. **Uyg'oq tutish**: GitHub repo → Settings → Secrets → Actions → `RENDER_URL` = `https://BOTNOMI.onrender.com`. Actions bo'limida "ping" ni yoqing (har 5 daqiqa). Zaxira: uptimerobot.com (bepul) ham xuddi shu manzilni tekshirishi mumkin.
7. Guruhlarga botni qo'shib `/id` yozing, chiqqan raqamlarni COLLECT_CHAT_ID / LOAD_CHAT_ID ga yozing.

## Bepul tarif cheklovlari
- Render: 750 soat/oy — faqat bitta bot doimiy ishlaydi; boshqa bepul servis qo'shmang.
- Supabase: 1 hafta faol bo'lmasa to'xtaydi (bot /health orqali bazani har 5 daqiqada "uyg'otib" turadi); avtomatik zaxira yo'q → vaqti-vaqti bilan botga `/zaxira` yozing (JSON fayl beradi).
- GitHub ping: repoda 60 kun harakat bo'lmasa o'chadi — Actions'da qayta yoqing.
