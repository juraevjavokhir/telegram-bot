# Заявка для Китая — Telegram bot

The bot runs in Docker on the Hetzner server. Every `git push` to `main`
deploys it automatically through GitHub Actions.

## One-time server setup (Hetzner)

```bash
ssh root@SERVER_IP

# Docker
curl -fsSL https://get.docker.com | sh

# Bot folder and token
mkdir -p /opt/telegram-bot
echo 'BOT_TOKEN=<token from @BotFather>' > /opt/telegram-bot/.env
chmod 600 /opt/telegram-bot/.env

# SSH key for GitHub Actions
ssh-keygen -t ed25519 -f ~/.ssh/gh_deploy -N ""
cat ~/.ssh/gh_deploy.pub >> ~/.ssh/authorized_keys
cat ~/.ssh/gh_deploy        # -> GitHub secret SSH_KEY
```

## GitHub secrets

Repo → Settings → Secrets and variables → Actions:

| Name       | Value                         |
|------------|-------------------------------|
| `SSH_HOST` | server IP address             |
| `SSH_USER` | `root`                        |
| `SSH_KEY`  | contents of `~/.ssh/gh_deploy` |

## Usage

- Deploy: `git push` (or Actions → Deploy → Run workflow)
- Logs: `ssh root@SERVER_IP "cd /opt/telegram-bot && docker compose logs -f"`
- Restart: `docker compose restart`

The token is stored only in `.env` on the server and is never committed to git.
