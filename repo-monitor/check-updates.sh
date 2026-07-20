#!/bin/bash
set -euo pipefail

# ===================== НАСТРОЙКИ =====================
WATCH_PACKAGES="openssh-server kernel"
LOCAL_REPO_URL="http://repo.corp.local/redos/8/x86_64/"  # ваш локальный URL
STATE_FILE="state.txt"

# =====================================================

# Проверяем, что API_TOKEN задан
if [ -z "${API_TOKEN:-}" ]; then
  echo "ОШИБКА: переменная API_TOKEN не установлена"
  exit 1
fi

GITLAB_API="https://gitlab.corp.com/api/v4/projects/${CI_PROJECT_ID}/issues"

for pkg in $WATCH_PACKAGES; do
  # Получаем последнюю версию пакета из локального репозитория
  current=$(dnf repoquery \
              --repofrompath=localrepo,${LOCAL_REPO_URL} \
              --repoid=localrepo \
              --latest-limit 1 \
              --qf "%{name}-%{epoch}:%{version}-%{release}.%{arch}" \
              "$pkg" 2>/dev/null || true)

  if [ -z "$current" ]; then
    echo "⚠️  Пакет '$pkg' не найден в $LOCAL_REPO_URL"
    continue
  fi

  # Читаем предыдущую версию из state.txt
  previous=$(grep "^${pkg} " "$STATE_FILE" 2>/dev/null | awk '{print $2}' || true)

  if [ "$previous" != "$current" ]; then
    echo "✅ Обнаружено обновление: $pkg  $previous → $current"

    # Создаём Issue в этом же проекте
    curl -s --request POST \
      --header "PRIVATE-TOKEN: ${API_TOKEN}" \
      --header "Content-Type: application/json" \
      --data "{
        \"title\": \"Обновление ${pkg} (${CI_PIPELINE_CREATED_AT%T*})\",
        \"description\": \"Пакет **${pkg}** обновлён в локальном репозитории.\n\n\`\`\`\nБыло : ${previous}\nСтало: ${current}\n\`\`\`\",
        \"labels\": [\"repo-monitoring\", \"auto\"]
      }" \
      "$GITLAB_API"

    # Обновляем state.txt
    if grep -q "^${pkg} " "$STATE_FILE" 2>/dev/null; then
      sed -i "s|^${pkg} .*|${pkg} ${current}|" "$STATE_FILE"
    else
      echo "${pkg} ${current}" >> "$STATE_FILE"
    fi
  else
    echo "Пакет $pkg не изменился ($current)"
  fi
done

# Выводим diff для информации (не влияет на результат)
git diff --exit-code "$STATE_FILE" || true