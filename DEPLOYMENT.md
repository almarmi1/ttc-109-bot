# Deploy to Coolify

1. Upload or push the updated project, including `github_assets.py`. Keep Excel files, images and PDF folders in the configured GitHub repository.
2. Select the **Dockerfile** build pack, using `/Dockerfile` in the repository root. The image starts `python TTC_109_v3.py`.
3. Add runtime environment variables in Coolify:
   - `BOT_TOKEN`: your Telegram bot token from BotFather.
   - `OWNER_ID`: your Telegram user ID.
   - `USERS_FILE=/data/users.json`.
4. Add persistent storage mounted at `/data` so registered users survive redeploys. To retain existing users, copy your existing users.json into that volume before starting. Do not mount an empty volume over `/app`, which contains the program and assets.
5. This is a long-polling background bot. Leave Domains empty, do not publish an HTTP port, and disable HTTP health checks. It does not serve a website.
6. Stop the local bot and any old deployments using the same token. Run only one polling instance.
7. Deploy and inspect runtime logs. Expect `Starting Telegram bot with long polling...`. Send `/start` to your bot in Telegram to verify it responds.

The container needs outbound HTTPS access to api.telegram.org and api.github.com. No inbound Telegram webhook port is required.

## Common failures

- `FileNotFoundError`: confirm all assets are committed/uploaded, with the exact filenames and folder names. Asset paths resolve in the configured GitHub repository and branch.
- Missing `BOT_TOKEN` or `InvalidToken`: set the runtime variable and redeploy. Local `.env` is deliberately excluded from Docker builds.
- `Conflict` / another getUpdates request: stop the other process using the bot token.
- `NetworkError` / `TimedOut`: check the VPS DNS and outbound HTTPS connectivity to Telegram.
- HTTP health check fails: disable it; this process has no HTTP server.
- `No matching distribution found`: inspect the build log, package index and Python compatibility. The Dockerfile uses Python 3.11.

The old source contained a bot token. Replace that token through BotFather and set the replacement in Coolify and your local .env; removing it from the current source does not remove it from Git history or earlier images.

## Read files from GitHub when selected

Deploy both `TTC_109_v3.py` and `github_assets.py`. The bot now fetches PDFs,
images, and Excel data from GitHub on demand instead of reading local assets.
Set `GITHUB_REPOSITORY=almarmi1/ttc-109-bot` and `GITHUB_REF=main` in Coolify
(or local `.env`). These are also the defaults. For a private repository, set
`GITHUB_TOKEN` to a token with read access to repository Contents. Keep the token
in runtime secrets, never in Git. Authenticated requests also have higher API limits.

Keep the existing asset filenames and folders in GitHub exactly as referenced,
including spaces and Arabic spelling. Files must be committed to the selected
branch; local uncommitted edits cannot be downloaded. Git LFS pointers are not
supported. Each file must be under 49 MiB. Allow outbound HTTPS to api.github.com.

Selecting a PDF/image downloads and sends that file. Staff, mentor, FAQ, and quiz
selections download the relevant Excel workbook. Staff lists, FAQ answers and quiz
questions remain consistent within the current selection/session; opening the
feature again fetches the latest version. No download occurs at bot startup.
Failures show an Arabic retry message and do not silently use an outdated local copy.

API reference: https://docs.github.com/en/rest/repos/contents#get-repository-content
