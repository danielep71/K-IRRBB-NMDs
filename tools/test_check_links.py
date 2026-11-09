"""Tests for the Markdown link checker (tools/check_links.py)."""

from __future__ import annotations

import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_links  # noqa: E402


def run_check(files: dict[str, str]) -> tuple[list[str], set[str], int]:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        for name, text in files.items():
            path = root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        # Not a Git checkout: the checker falls back to every *.md file.
        return check_links.check(root)


class SlugTests(unittest.TestCase):
    def test_github_slugs(self) -> None:
        self.assertEqual(check_links.github_slug("Branches and publication"), "branches-and-publication")
        self.assertEqual(check_links.github_slug("🧭 Before you start"), "-before-you-start")
        # U+FE0F after an emoji is a combining mark that GitHub keeps; "①" is not a word character.
        self.assertEqual(check_links.github_slug("▶\ufe0f Launch sequence"), "\ufe0f-launch-sequence")
        self.assertEqual(check_links.github_slug("① Preview the changes"), "-preview-the-changes")
        self.assertEqual(check_links.github_slug("Café 2026_v1"), "café-2026_v1")
        self.assertEqual(check_links.github_slug("Run `check.py` [now](x.md)"), "run-checkpy-now")
        self.assertEqual(check_links.github_slug("Rock &amp; Roll"), "rock--roll")

    def test_duplicate_and_setext_headings(self) -> None:
        _, anchors = check_links.markdown_links("# Notes\n\n## Notes\n\nTitle\n=====\n")
        self.assertEqual(anchors, {"notes", "notes-1", "title"})
        _, anchors = check_links.markdown_links("# Foo\n# Foo-1\n# Foo\n")
        self.assertEqual(anchors, {"foo", "foo-1", "foo-2"})

    def test_code_is_ignored_and_explicit_anchors_count(self) -> None:
        text = ('<a id="custom"></a>\n\n```sh\n[x](missing.md)\n# not a heading\n```\n'
                "`[y](also-missing.md)` [z](real.md#part)\n")
        links, anchors = check_links.markdown_links(text)
        self.assertEqual(links, ["real.md#part"])
        self.assertEqual(anchors, {"custom"})

    def test_indented_code_blocks_are_ignored(self) -> None:
        text = ("Intro\n\n    [a](missing.md)\n    # not a heading\n\n\t[b](missing.md)\n\n"
                "# Real\n    [c](missing.md)\n\nText [d](real.md)\n")
        links, anchors = check_links.markdown_links(text)
        self.assertEqual(links, ["real.md"])
        self.assertEqual(anchors, {"real"})

    def test_indentation_that_is_not_code_keeps_its_links(self) -> None:
        # A paragraph continuation and list-item content are not code blocks.
        text = ("Paragraph\n    [a](a.md)\n\n- item\n\n    [b](b.md)\n"
                "  1. nested\n\n        [c](c.md)\n")
        links, _ = check_links.markdown_links(text)
        self.assertEqual(links, ["a.md", "b.md", "c.md"])

    def test_a_paragraph_after_a_list_ends_it(self) -> None:
        links, _ = check_links.markdown_links("- item\n\nText\n\n    [a](missing.md)\n")
        self.assertEqual(links, [])


class CheckTests(unittest.TestCase):
    def test_valid_links_pass_and_external_links_are_collected(self) -> None:
        errors, external, count = run_check({
            "README.md": "[Guide](docs/guide.md#-install) · [site](https://example.com/page#top)\n"
                         "[self](#readme) <https://example.org>\n\n# README\n",
            "docs/guide.md": "# 🚀 Install\n\n[Back](../README.md)\n",
        })
        self.assertEqual(errors, [])
        self.assertEqual(external, {"https://example.com/page", "https://example.org"})
        self.assertEqual(count, 5)

    def test_missing_files_anchors_and_paths_are_errors(self) -> None:
        errors, _, _ = run_check({
            "README.md": "[a](docs/missing.md) [b](docs/guide.md#nowhere) [c](/abs.md)\n"
                         "[d](../outside.md) [e](empty.md) [f](ftp://example.com)\n",
            "docs/guide.md": "# Guide\n",
            "empty.md": "",
        })
        self.assertEqual(errors, [
            "README.md: missing target: docs/missing.md",
            "README.md: missing anchor: docs/guide.md#nowhere",
            "README.md: root-relative link: /abs.md",
            "README.md: path outside the repository: ../outside.md",
            "README.md: empty target: empty.md",
            "README.md: unsupported URL scheme: ftp://example.com",
        ])

    def test_plain_http_links_are_errors_outside_code(self) -> None:
        errors, external, count = run_check({
            "README.md": "[a](http://a.example/x) <http://b.example> [c](HTTP://c.example)\n"
                         "[d]: http://d.example/\n<a href=\"http://e.example\">e</a>\n"
                         "`[f](http://f.example)` and `<http://g.example>` are code\n\n"
                         "```text\n[h](http://h.example)\n```\n"
                         "[ok](https://ok.example) [relative](//cdn.example/x#top)\n",
        })
        self.assertEqual(errors, [
            "README.md: insecure link (use https): http://a.example/x",
            "README.md: insecure link (use https): HTTP://c.example",
            "README.md: insecure link (use https): http://b.example",
            "README.md: insecure link (use https): http://d.example/",
            "README.md: insecure link (use https): http://e.example",
        ])
        self.assertEqual(external, {"https://ok.example", "https://cdn.example/x"})
        self.assertEqual(count, 7)

    def test_tracked_file_missing_from_the_working_tree_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text("[guide](guide.md)\n", encoding="utf-8")
            (root / "guide.md").write_text("# Guide\n", encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            (root / "guide.md").unlink()
            errors, _, _ = check_links.check(root)
        self.assertEqual(errors, ["guide.md: tracked but missing from the working tree",
                                  "README.md: missing target: guide.md"])

    def test_repository_links_pass(self) -> None:
        errors, _, count = check_links.check(ROOT)
        self.assertGreater(count, 0)
        self.assertEqual(errors, [])


class ExternalTests(unittest.TestCase):
    def test_own_repository_urls_are_skipped(self) -> None:
        own = check_links.own_repository
        self.assertTrue(own("https://github.com/Owner/Repo/blob/HEAD/x.md", "owner/repo"))
        self.assertTrue(own("https://github.com/owner/repo", "owner/repo"))
        self.assertTrue(own("https://raw.githubusercontent.com/owner/repo/main/x", "owner/repo"))
        self.assertFalse(own("https://github.com/owner/repo-other", "owner/repo"))
        self.assertFalse(own("https://github.com/owner/repo", ""))

    def test_only_errors_fail_and_own_urls_are_not_requested(self) -> None:
        results = {"https://a.example/": ("PASS", "https://a.example/"),
                   "https://b.example/": ("WARN", "https://b.example/: unavailable"),
                   "https://c.example/": ("ERROR", "https://c.example/: HTTP 404")}
        with unittest.mock.patch.object(check_links, "check_external",
                                        side_effect=lambda url: results[url]) as fake, \
                contextlib.redirect_stdout(io.StringIO()):
            errors = check_links.external_errors(
                [*results, "https://github.com/owner/repo/issues"], "owner/repo")
        self.assertEqual(errors, ["https://c.example/: HTTP 404"])
        self.assertEqual(sorted(call.args[0] for call in fake.call_args_list), sorted(results))

    def test_http_codes_are_classified(self) -> None:
        def curl(code: str, returncode: int = 0) -> unittest.mock.Mock:
            return unittest.mock.Mock(stdout=code, returncode=returncode)

        for code, returncode, status in (("200", 0, "PASS"), ("404", 0, "ERROR"),
                                         ("429", 0, "WARN"), ("503", 0, "WARN"),
                                         ("", 6, "WARN"), ("301", 1, "ERROR"),
                                         ("000", 1, "WARN")):
            with self.subTest(code=code), unittest.mock.patch(
                    "check_links.subprocess.run", return_value=curl(code, returncode)):
                self.assertEqual(check_links.check_external("https://x.example/")[0], status)

    def test_requests_use_https_only(self) -> None:
        with unittest.mock.patch("check_links.subprocess.run",
                                 return_value=unittest.mock.Mock(stdout="200", returncode=0)) as run:
            check_links.check_external("https://x.example/")
        command = run.call_args.args[0]
        self.assertEqual(command[command.index("--proto") + 1], "=https")
        self.assertEqual(command[command.index("--proto-redir") + 1], "=https")

    def test_a_refused_redirect_is_an_error(self) -> None:
        # curl stops with exit status 1 (unsupported protocol) after the 3xx answer.
        refused = unittest.mock.Mock(
            stdout="302", returncode=1,
            stderr='curl: (1) Protocol "http" not supported or disabled in libcurl\n')
        with unittest.mock.patch("check_links.subprocess.run", return_value=refused):
            self.assertEqual(check_links.check_external("https://x.example/"),
                             ("ERROR", "https://x.example/: redirects to a URL that is not HTTPS "
                                       "(HTTP 302)"))


if __name__ == "__main__":
    unittest.main()
