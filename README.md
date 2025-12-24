# qwertyu — ProjectAI

## Что есть
- Генератор: тема → пошаговая генерация по этапам → экспорт DOCX.
- Панель управления:
  - Пресеты: создать/удалить/сделать активным.
  - Вкладка «Этапы»: включать/выключать, менять порядок, добавлять/удалять этапы, управлять количеством глав и параграфов.
  - Вкладка «Промпты»:
    - редактирование любых промптов,
    - создание новых promptId,
    - дублирование,
    - удаление,
    - переименование promptId с безопасным обновлением всех этапов и defaultPromptIds.
  - Переменные + авто-контекст (summary/tail).

## Важно
AI НЕ работает при открытии index.html как file://  
Нужно запускать через Vercel.

## Локальный запуск (Windows PowerShell)
1) Установите Node.js (LTS): https://nodejs.org  
2) Установите Vercel CLI:
   npm i -g vercel
3) В папке проекта:
   vercel login
   vercel env add DEEPSEEK_API_KEY
   vercel dev

Откройте:
- http://localhost:3000 — генератор
- http://localhost:3000/prompts.html — панель управления

## Примечание
Если вы переименовали ключевой promptId (например, section), синхронизация этапов глав будет использовать обновлённый defaultPromptIds.section.


## Запуск локально (Windows / PowerShell)

1) Перейдите в папку проекта:
```powershell
cd "C:\Users\Admin\Desktop\пушка\v9t"
```

2) Установите зависимости (обязательно для серверного экспорта DOCX через npm-пакет `docx`):
```powershell
npm install
```

3) Укажите API-ключ DeepSeek в `.env.local` (пример, подставьте свой реальный ключ):
```powershell
'DEEPSEEK_API_KEY="sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"' | Out-File -Encoding utf8 .env.local
```

4) Запустите dev-сервер:
```powershell
vercel dev
```

5) Откройте в браузере:
- http://localhost:3000

Если экспорт DOCX не работает, откройте DevTools → Network и посмотрите ответ `/api/docx`.
