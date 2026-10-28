# Deploy to Coolify

1. Upload or push the updated project, including all Excel files, images and PDF folders.
2. Select the **Dockerfile** build pack, using `/Dockerfile` in the repository root. The image starts `python TTC_109_v3.py`.
3. Add runtime environment variables in Coolify:
   - `BOT_TOKEN`: your Telegram bot token from BotFather.
   - `OWNER_ID`: your Telegram user ID.
   - `USERS_FILE=/data/users.json`.
4. Add persistent storage mounted at `/data` so registered users survive redeploys. To retain existing users, copy your existing users.json into that volume before starting. Do not mount an empty volume over `/app`, which contains the program and assets.
5. This is a long-polling background bot. Leave Domains empty, do not publish an HTTP port, and disable HTTP health checks. It does not serve a website.
6. Stop the local bot and any old deployments using the same token. Run only one polling instance.
7. Deploy and inspect runtime logs. Expect `Starting Telegram bot with long polling...`. Send `/start` to your bot in Telegram to verify it responds.

The container needs outbound HTTPS access to api.telegram.org. No inbound Telegram webhook port is required.

## Common failures

- `FileNotFoundError`: confirm all assets are committed/uploaded, with the exact filenames and folder names. Paths now resolve beside the Python script.
- Missing `BOT_TOKEN` or `InvalidToken`: set the runtime variable and redeploy. Local `.env` is deliberately excluded from Docker builds.
- `Conflict` / another getUpdates request: stop the other process using the bot token.
- `NetworkError` / `TimedOut`: check the VPS DNS and outbound HTTPS connectivity to Telegram.
- HTTP health check fails: disable it; this process has no HTTP server.
- `No matching distribution found`: inspect the build log, package index and Python compatibility. The Dockerfile uses Python 3.11.

The old source contained a bot token. Replace that token through BotFather and set the replacement in Coolify and your local .env; removing it from the current source does not remove it from Git history or earlier images.
