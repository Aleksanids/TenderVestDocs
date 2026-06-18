# Stage lifecycle

## Состояния этапа

```text
draft -> implementation -> local validation -> PR -> CI audit -> AI review -> user acceptance -> merge -> post-merge note
```

## draft

Фиксируются stage, цель, scope, запреты, ожидаемые проверки и критерии готовности.

## implementation

Codex делает минимальные изменения в отдельной ветке. Изменения вне scope не допускаются.

## local validation

Запускаются доступные локальные проверки. Если проверка не запускается, причина фиксируется в отчете.

## PR

Открывается Pull Request с заполненным шаблоном и ссылками на отчеты.

## CI audit

GitHub Actions запускает safe/offline audit: compileall, pytest и safety scan запрещенных файлов.

## AI review

ChatGPT и Claude проверяют diff, отчеты, tests, warnings, limitations и file-safety.

## user acceptance

Пользователь принимает решение: `ok`, `ok_with_warnings`, `needs_hotfix` или `rejected`.

## merge

Merge разрешен только после приемки. Branch protection должен запрещать прямые изменения `main`.

## post-merge note

После merge фиксируется краткая запись: что вошло, какие warnings остались, какой следующий этап рекомендован.
