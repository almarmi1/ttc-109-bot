"""Fetch repository assets on demand; no local asset files are required."""
from io import BytesIO
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen


class GitHubAssets:
    MAX_BYTES = 49 * 1024 * 1024

    def __init__(self, repository, ref="main", token=""):
        self.repository = repository.strip()
        if len(self.repository.split("/")) != 2 or any(
            not part or part in (".", "..") for part in self.repository.split("/")
        ):
            raise ValueError("GITHUB_REPOSITORY must be owner/repository")
        self.ref = ref.strip() or "main"
        self.token = token.strip()

    def read(self, path):
        if path.startswith("/") or any(part in (".", "..") for part in path.split("/")):
            raise ValueError("Asset path must be relative to the repository")
        # Encode each exact filename; do not strip spaces or normalize Unicode.
        url = (
            "https://api.github.com/repos/" + quote(self.repository, safe="/")
            + "/contents/" + quote(path, safe="/")
            + "?" + urlencode({"ref": self.ref})
        )
        headers = {
            "Accept": "application/vnd.github.raw+json",
            "User-Agent": "ttc-109-bot",
        }
        if self.token:
            headers["Authorization"] = "Bearer " + self.token
        with urlopen(Request(url, headers=headers), timeout=30) as response:
            data = response.read(self.MAX_BYTES + 1)
        if len(data) > self.MAX_BYTES:
            raise ValueError("Asset exceeds the bot's 49 MiB upload limit")
        if data.startswith(b"version https://git-lfs.github.com/spec/v1"):
            raise ValueError("Git LFS pointers are not supported; commit the asset directly")
        result = BytesIO(data)
        result.name = path.rsplit("/", 1)[-1]
        return result
