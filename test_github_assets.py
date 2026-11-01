import asyncio
import importlib
import unittest
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch
from urllib.error import HTTPError
from urllib.parse import unquote
from github_assets import GitHubAssets


class AssetTests(unittest.TestCase):
    @patch('github_assets.urlopen')
    def test_exact_arabic_path_and_private_auth(self, fetch):
        fetch.return_value.__enter__.return_value.read.return_value = b'%PDF-data'
        path = 'قسم الدراسات العامة/ الدراسات الإسلامية.pdf'
        result = GitHubAssets('owner/repo', 'feature/files', 'secret').read(path)
        request = fetch.call_args.args[0]
        self.assertIn(path, unquote(request.full_url))
        self.assertIn('ref=feature%2Ffiles', request.full_url)
        self.assertEqual(request.get_header('Authorization'), 'Bearer secret')
        self.assertEqual(result.read(), b'%PDF-data')
        self.assertEqual(result.name, path.split('/')[-1])

    @patch('github_assets.urlopen')
    def test_remote_error_is_not_local_fallback(self, fetch):
        fetch.side_effect = HTTPError('url', 404, 'missing', {}, None)
        with self.assertRaises(HTTPError):
            GitHubAssets('owner/repo').read('missing.pdf')

    @patch('github_assets.urlopen')
    def test_size_and_lfs_rejected(self, fetch):
        client = GitHubAssets('owner/repo')
        client.MAX_BYTES = 2
        fetch.return_value.__enter__.return_value.read.return_value = b'abc'
        with self.assertRaises(ValueError):
            client.read('large.pdf')
        client.MAX_BYTES = 1000
        fetch.return_value.__enter__.return_value.read.return_value = b'version https://git-lfs.github.com/spec/v1'
        with self.assertRaises(ValueError):
            client.read('pointer.pdf')

    @patch('github_assets.urlopen')
    def test_startup_and_selection(self, fetch):
        bot = importlib.import_module('TTC_109_v3')
        fetch.assert_not_called()
        message = SimpleNamespace(reply_document=AsyncMock(), reply_photo=AsyncMock(), reply_text=AsyncMock())
        update = SimpleNamespace(message=message)
        async def run():
            context = SimpleNamespace(user_data={'state': 'trainee_guides'})
            message.text = next(iter(bot.trainee_guides_files))
            with patch.object(bot.assets, 'read', return_value=BytesIO(b'%PDF-test')) as read:
                await bot.handle_choice(update, context)
                read.assert_called_once_with(bot.trainee_guides_files[message.text])
                message.reply_document.assert_awaited_once()
            with patch.object(bot.assets, 'read', side_effect=TimeoutError):
                await bot.send_asset(update, 'missing.pdf')
                message.reply_text.assert_awaited_once()
        asyncio.run(run())


if __name__ == '__main__':
    unittest.main()
