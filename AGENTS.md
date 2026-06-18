# AGENTS.md - TenderVestDocs GitHub workflow

Всегда отвечай пользователю на русском языке. Все пользовательские тексты, Markdown, JSON, README, отчеты и статусы пиши в UTF-8.

## Базовые правила

- Один этап = одна ограниченная задача с явно описанным scope.
- Не работать напрямую в `main`.
- Не создавать новый clone без прямой команды пользователя.
- Не менять исходные пользовательские Excel, CSV, DOCX, PDF, ZIP, HTML и другие документы.
- Не запускать live network, ЕИС, ЭТП, browser automation, Selenium или Playwright без прямого разрешения.
- Не читать cookies, tokens, passwords, browser profiles, `.env` и другие секреты.
- Не создавать EXE, installer, PyInstaller build или иной build-артефакт без прямой команды.
- Не добавлять пользовательские документы, скачанные документы закупок, runtime/cache/temp/downloads и тяжелые архивы в GitHub.
- Все результаты этапа оформлять через отдельную ветку и Pull Request.

## Отчетность этапа

Все GitHub-отчеты этапа складывать в:

```text
reports/stages/<STAGE_NAME>/
```

Локальные handoff-материалы складывать в соответствующую папку передачи:

```text
D:\Codex\TenderVestDocs\Передать в GPT\<STAGE>
D:\Codex\TenderVestDocs\00_Передать_Claude\01_ОТПРАВИТЬ_В_CLAUDE\<STAGE>
```

Для текущего GitHub-пакета используется локальная рабочая папка:

```text
D:\Codex\TenderVestDocs\00_Передать_ChatGPT\git
```

## Финальный отчет Codex

В финальном отчете всегда указывать:

- status;
- changed files;
- created files;
- commands run;
- tests;
- not run;
- warnings;
- limitations;
- safety confirmations.

## Safety confirmations

Перед завершением этапа подтвердить:

- исходные пользовательские Excel / CSV / DOCX / PDF не изменялись;
- пользовательские документы не добавлялись в репозиторий;
- скачанные документы не добавлялись в репозиторий;
- runtime / cache / temp / downloads не добавлялись;
- cookies / tokens / passwords / browser profiles не читались;
- live network не запускался без прямого разрешения;
- browser automation не запускалась без прямого разрешения;
- EXE / build не создавался без прямой команды;
- изменения ограничены scope этапа.
