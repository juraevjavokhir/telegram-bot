# Заявка для Китая — Telegram bot

The bot collects a list of spare parts and builds an Excel request
(№, Фото, Наименование, Артикул, Количество, Ед. изм., Цена, Сумма) for the
Chinese partner.

## How to add items quickly

| What you send | Result |
|---|---|
| Photo with caption `Полуось` ↵ `A3463530435` ↵ `2` | item added in one message |
| Text, one item per line: `Полуось; A3463530435; 2; шт` (Excel rows can be pasted too) | many items at once |
| `A3463530435 2` (article already ordered before) | name/unit filled from the catalog |
| Album of photos without captions | bot asks only the missing fields for each |
| An old request `.xlsx` as a file | all rows (with photos) loaded |
| Photo without caption | asks name → article (`-` = none) → quantity |

Commands: `/list`, `/qty 5 10` (set qty of №5), `/del 5`, `/skip`, `/done`,
`/cancel`, `/help`. Default unit is `шт`; units are normalized (`Шт`, `ш` → `шт`);
the same article added twice is merged (quantities summed).

Unfinished requests, photos and the article catalog are stored in `./data`
on the server (Docker volume), so a restart does not lose them.

## Deploy

The bot runs in Docker on the Hetzner server (`/opt/telegram-bot`). Every
`git push` to `main` deploys it through GitHub Actions
(`.github/workflows/deploy.yml`): files are copied over SSH as user `deploy`
and `docker compose up -d --build` is run. A new file the bot needs must be
added to the workflow's `source:` list.

GitHub secrets: `SSH_HOST` (server IP), `SSH_USER` (`deploy`), `SSH_KEY`
(private key `~/.ssh/gh_actions` from the server, authorized for `deploy`).
Root SSH login is disabled on the server.

- Logs: `ssh deploy@SERVER_IP "cd /opt/telegram-bot && docker compose logs --tail 50 bot"`
- Restart: `ssh deploy@SERVER_IP "cd /opt/telegram-bot && docker compose restart"`

The token is stored only in `/opt/telegram-bot/.env` on the server
(`BOT_TOKEN=...`) and is never committed to git.
