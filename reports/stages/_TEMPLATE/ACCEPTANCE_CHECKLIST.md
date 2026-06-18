# Acceptance checklist

## Scope

- [ ] PR соответствует заявленному stage.
- [ ] Изменены только разрешенные файлы.
- [ ] Нет изменений в коде приложения вне scope.

## Reports

- [ ] Есть stage report.
- [ ] Есть tests report.
- [ ] Есть manifest.
- [ ] Есть handoff path.

## Checks

- [ ] GitHub Actions `codex_audit` завершился успешно или предупреждения объяснены.
- [ ] Локальные проверки указаны в отчете.
- [ ] Незапущенные проверки явно перечислены.

## File safety

- [ ] Пользовательские документы не добавлены.
- [ ] Runtime / cache / temp / downloads не добавлены.
- [ ] Secrets не добавлены.
- [ ] Live network не запускался без разрешения.
- [ ] Browser automation не запускалась без разрешения.
- [ ] EXE / build не создавался.

## Decision

- [ ] ok
- [ ] ok_with_warnings
- [ ] needs_hotfix
- [ ] rejected
