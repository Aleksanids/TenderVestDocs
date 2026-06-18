# Recommended branch protection

## Цель

Защитить `main` от прямых изменений и сделать Pull Request обязательной точкой приемки.

## Рекомендуемые настройки для `main`

- Require a pull request before merging.
- Require status checks before merging.
- Require successful `codex_audit` workflow.
- Disallow force pushes.
- Disallow deletions.
- Require linear history: optional.
- Require review before merging: optional.
- Include administrators: optional.

## Команды для ручной настройки

Не применять автоматически без отдельного разрешения пользователя.

```bash
gh api \
  --method PUT \
  repos/Aleksanids/TenderVestDocs/branches/main/protection \
  --field required_status_checks='{"strict":true,"contexts":["codex_audit"]}' \
  --field enforce_admins=true \
  --field required_pull_request_reviews='{"required_approving_review_count":1}' \
  --field restrictions=null \
  --field allow_force_pushes=false \
  --field allow_deletions=false
```

## Ограничение

Настройки branch protection не меняются в рамках `GITHUB_WORKFLOW_01` без отдельной прямой команды.
