# GitHub workflow TenderVestDocs

## Цель

GitHub используется как центр учета этапов, отчетности Codex, хранения diff, запуска safe/offline проверок и приемки изменений.

## Полный цикл

1. Codex получает prompt и фиксирует stage, цель, scope и запреты.
2. Codex создает отдельную ветку от `main`.
3. Codex делает минимальные изменения в разрешенной области.
4. Codex запускает локальные проверки.
5. Codex создает отчет этапа в `reports/stages/<STAGE_NAME>/`.
6. Codex открывает Pull Request.
7. GitHub Actions выполняет `codex_audit`.
8. Claude / ChatGPT проводят review по своим ролям.
9. Пользователь принимает решение.
10. Merge разрешен только после статуса `ok` или `ok_with_warnings`.

## Правила веток

- Не работать напрямую в `main`.
- Один stage = одна ветка.
- Имя ветки должно отражать stage или hotfix.
- Если рабочее дерево содержит чужие изменения, нельзя добавлять их в PR без явного подтверждения.

## Правила PR

Pull Request должен содержать:

- stage;
- цель;
- scope;
- changed files;
- created files;
- commands run;
- tests;
- not run;
- handoff path;
- warnings and limitations;
- safety confirmations;
- next step.

## Что нельзя добавлять в GitHub

- исходные пользовательские Excel / CSV / DOCX / PDF;
- скачанные документы закупок;
- runtime / cache / temp / downloads;
- AI-ready архивы;
- cookies / tokens / passwords / browser profiles;
- EXE / installer / build outputs.
