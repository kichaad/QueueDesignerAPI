# Очередь заказов

Google Apps Script V8 project imported from the owner's local project. Source is in `src/`; `.clasp.json` identifies the existing Google script. This replaces the old root-level layout on this branch. No application logic was changed during import.

## Development

Node.js >=20; run `npm ci`, `node scripts/check-source.cjs`, and `npm run status`. There is no local web server. Syntax checks do not validate Google services or business behavior. Do not run `npm run pull` over unsaved local changes.

## GitHub publication setup

1. Enable Google Apps Script API at https://script.google.com/home/usersettings for the account owning this script.
2. On the trusted local computer run `npm run login` if needed. Never upload `.clasprc.json` to chat or Git.
3. In this repository open Settings → Environments and create `apps-script`. Add an environment secret `CLASP_AUTH_JSON`: securely paste the entire contents of the local `.clasprc.json` directly into GitHub's secret form. Prefer a dedicated authorization; revocation or OAuth expiration requires renewal. Enable environment deployment reviewers where available.
4. Add an environment variable `GAS_DEPLOYMENT_ID` for the existing deployment. Verify the ID in Apps Script before setting it; do not use the scriptId here.
5. Merge this reviewed branch into the default branch. In Actions select **Apps Script**, **Run workflow**, and operation **verify** first. This only lists deployments. GitHub normally requires a workflow to exist on the default branch before manual dispatch is available.
6. After checking the remote source against this imported snapshot, run operation **publish** on the intended branch/commit. It overwrites remote editor source with `src/`, creates an immutable version, and updates the specified existing deployment while preserving its URL. Do not publish until the owner has confirmed this snapshot is current. No publication occurs on push or pull requests.

The credential exists only in a restricted temporary file on the Actions runner and is removed at job exit. It is not stored in the Codex cloud snapshot. Logs must never print its contents. Google OAuth, workflow permissions, and actual publication have not yet been verified. Cloud Codex needs separate GitHub API/Actions authorization to trigger jobs itself; Git read/push access does not establish that capability.

The configured spreadsheet ID is preserved from the source. The purpose of the second spreadsheet provided during handoff remains unconfirmed. Existing Script Properties remain in the same Apps Script project. Real application calls can write to Sheets/Drive or send notifications; use a separate test setup before functional testing.

## Compare published version

Run Actions → Apps Script → Run workflow → operation `compare`, branch `master`. This reads the exact immutable version referenced by `GAS_DEPLOYMENT_ID` and compares it to the selected repository commit, without pushing or publishing. The job summary lists identical counts and changed, repository-only, and published-only file names. Only filenames and counts are logged; source and OAuth values are not printed. This does not compare Script Properties, triggers, spreadsheet contents or runtime behavior. Comparison code uses the installed clasp authentication adapter from locked clasp 3.4.1; revalidate it when upgrading clasp.


## МДМ: CSV корзины

В «Снабжение → Заказы поставщикам» кнопка «Выгрузить МДМ CSV» отбирает из текущего списка только МДМ / «Счёт получен». Галочки выбирают целые заказы: частичный заказ не помечается выгруженным. Одинаковые артикулы объединяются, количество суммируется (единицы должны совпадать). Формат: UTF-8 без BOM и заголовка, `артикул;количество`, CRLF, без конечной точки с запятой; внутренние пробелы артикула сохраняются, крайние удаляются, десятичный разделитель — запятая. Максимум 500 уникальных артикулов; некорректные данные блокируют формирование без смены статусов.

После подготовки нажмите «Сохранить CSV…». Если браузер разрешает выбор файла, выберите папку. Apps Script работает в iframe и может запрещать этот API: тогда используется обычное скачивание. Чтобы браузер спрашивал папку, включите «Всегда указывать место для скачивания» в его настройках загрузок. При обычном скачивании приложение не может узнать его результат: нажмите «Файл сохранён — отметить „Выгружен“» только после сохранения, при отмене закройте окно. При выборе файла через API статус меняется после успешного завершения записи. Файлы не загружаются в Drive или на сайт МДМ автоматически.

Подготовка действительна 30 минут (кэш может быть очищен раньше). Перед сменой статуса состав и реквизиты повторно сверяются; при изменениях нужно сформировать файл заново. Повторное подтверждение безопасно для того же состава. Выгрузка меняет статус заказов на «Выгружен», материалы остаются на этапе «Счёт получен». Проверка: `node tests/mdm-export.cjs`. Живой импорт в корзину МДМ и рабочие Google-таблицы автоматически не тестируются.
