"""Read back the development workbook's VBA project with an independent reader.

tools/devbuild/vba_project.py writes vbaProject.bin from the [MS-OVBA] and
[MS-CFB] specifications. These tests decompress and parse its output with a
separate, minimal implementation of the same specifications, so a writer defect
fails here without Excel. Standard library only; the workbook assembly itself
(openpyxl) is exercised only when openpyxl is installed.
"""
from __future__ import annotations

import io
import random
import struct
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
import vba_project  # noqa: E402

ENDOFCHAIN = 0xFFFFFFFE


def decompress(data: bytes) -> bytes:
    """[MS-OVBA] 2.4.1.3.1: decompress a CompressedContainer."""
    if not data or data[0] != 0x01:
        raise ValueError("missing signature byte")
    out = bytearray()
    pos = 1
    while pos < len(data):
        start = pos
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        if (header >> 12) & 0x7 != 0b011:
            raise ValueError("bad chunk signature")
        end = min(start + (header & 0x0FFF) + 3, len(data))
        if not header & 0x8000:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        chunk_start = len(out)
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                    continue
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                difference = len(out) - chunk_start
                bit_count = 4
                while (1 << bit_count) < difference:
                    bit_count += 1
                length_mask = 0xFFFF >> bit_count
                length = (token & length_mask) + 3
                offset = (token >> (16 - bit_count)) + 1
                source = len(out) - offset
                for i in range(length):
                    out.append(out[source + i])
    return bytes(out)


class CompoundFile:
    """[MS-CFB] version-3 reader: FAT, directory tree and mini stream."""

    def __init__(self, blob: bytes) -> None:
        if blob[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError("not a compound file")
        self.blob = blob
        self.sector = 1 << struct.unpack_from("<H", blob, 30)[0]
        self.mini = 1 << struct.unpack_from("<H", blob, 32)[0]
        n_fat, dir_start = struct.unpack_from("<II", blob, 44)
        self.cutoff, minifat_start = struct.unpack_from("<II", blob, 56)
        difat = struct.unpack_from("<109I", blob, 76)[:n_fat]
        fat_bytes = b"".join(self._sector(s) for s in difat)
        self.fat = list(struct.unpack(f"<{len(fat_bytes) // 4}I", fat_bytes))
        directory = self._chain_bytes(dir_start)
        self.entries = [self._entry(directory[i:i + 128]) for i in range(0, len(directory), 128)]
        root = self.entries[0]
        self.mini_stream = self._chain_bytes(root["start"])[:root["size"]]
        minifat_bytes = self._chain_bytes(minifat_start) if minifat_start != ENDOFCHAIN else b""
        self.minifat = list(struct.unpack(f"<{len(minifat_bytes) // 4}I", minifat_bytes))

    def _sector(self, index: int) -> bytes:
        offset = self.sector * (index + 1)
        return self.blob[offset:offset + self.sector]

    def _chain(self, start: int, table: list[int]) -> list[int]:
        chain, seen = [], set()
        while start != ENDOFCHAIN:
            if start in seen:
                raise ValueError("cyclic chain")
            seen.add(start)
            chain.append(start)
            start = table[start]
        return chain

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self._sector(s) for s in self._chain(start, self.fat))

    @staticmethod
    def _entry(raw: bytes) -> dict:
        name_length = struct.unpack_from("<H", raw, 64)[0]
        return {"name": raw[:max(0, name_length - 2)].decode("utf-16-le"), "kind": raw[66],
                "left": struct.unpack_from("<I", raw, 68)[0], "right": struct.unpack_from("<I", raw, 72)[0],
                "child": struct.unpack_from("<I", raw, 76)[0], "start": struct.unpack_from("<I", raw, 116)[0],
                "size": struct.unpack_from("<Q", raw, 120)[0]}

    def children(self, sid: int) -> dict[str, int]:
        found: dict[str, int] = {}
        stack = [self.entries[sid]["child"]]
        while stack:
            node = stack.pop()
            if node in (0xFFFFFFFF, None):
                continue
            entry = self.entries[node]
            found[entry["name"]] = node
            stack += [entry["left"], entry["right"]]
        return found

    def read(self, path: str) -> bytes:
        sid = 0
        for part in path.split("/"):
            sid = self.children(sid)[part]
        entry = self.entries[sid]
        if entry["size"] >= self.cutoff:
            return self._chain_bytes(entry["start"])[:entry["size"]]
        if entry["size"] == 0:
            return b""
        chunks = [self.mini_stream[s * self.mini:(s + 1) * self.mini] for s in self._chain(entry["start"], self.minifat)]
        return b"".join(chunks)[:entry["size"]]


def sample_modules() -> list[vba_project.Module]:
    body = "\r\n".join(f"' line {i}: some repeated text, repeated text, repeated text" for i in range(200))
    return [vba_project.document_module("ThisWorkbook", True),
            vba_project.document_module("shChecks", False),
            vba_project.Module("CORE_Small", 'Attribute VB_Name = "CORE_Small"\r\nOption Explicit\r\n'),
            vba_project.Module("TEST_Large", 'Attribute VB_Name = "TEST_Large"\r\n' + body + "\r\n")]


class Compression(unittest.TestCase):
    def test_specification_example_decompresses(self) -> None:
        # [MS-OVBA] 3.2.3, "Normal Compression" example.
        compressed = bytes.fromhex(
            "012FB000236161616263646582660070616768696A01380861"
            "6B6C00306D6E6F7006710270041072737475761077"
            "78797A00" "3C")
        self.assertEqual(decompress(compressed), b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa")

    def test_round_trip(self) -> None:
        rng = random.Random(8)
        samples = [b"", b"a", b"abc" * 3000, bytes(rng.randrange(256) for _ in range(9000)),
                   b"Option Explicit\r\n" * 700, bytes(range(256)) * 40]
        for data in samples:
            with self.subTest(length=len(data)):
                self.assertEqual(decompress(vba_project.compress(data)), data)


    def test_incompressible_last_chunk_names_the_module(self) -> None:
        # About 4 KB of random printable text in the last, short chunk does not fit
        # one compressed chunk; the build stops and names the module.
        rng = random.Random(97)
        noise = "".join(chr(rng.randrange(33, 127)) for _ in range(4000))
        module = vba_project.Module("CORE_Noise", f'Attribute VB_Name = "CORE_Noise"\r\n\' {noise}\r\n')
        with self.assertRaisesRegex(ValueError, r"module CORE_Noise: .*Manual import"):
            vba_project.build_vba_project([module], ProjectStorage.project_id)


class Encryption(unittest.TestCase):
    def test_project_protection_fields_round_trip(self) -> None:
        rng = random.Random("{ID}")
        for data in (b"\x00", b"\xff", struct.pack("<I", 0), b"abcdefgh"):
            self.assertEqual(vba_project.decrypt(vba_project._encrypt(data, "{ID}", rng)), data)


class ProjectStorage(unittest.TestCase):
    project_id = "{0B9A3C1D-1111-4222-8333-444455556666}"

    def build(self) -> CompoundFile:
        return CompoundFile(vba_project.build_vba_project(sample_modules(), self.project_id))

    def test_every_module_source_reads_back_unchanged(self) -> None:
        cfb = self.build()
        for module in sample_modules():
            with self.subTest(module=module.name):
                self.assertEqual(decompress(cfb.read(f"VBA/{module.name}")).decode("cp1252"), module.source)

    def test_project_and_dir_streams_list_every_module(self) -> None:
        cfb = self.build()
        project = cfb.read("PROJECT").decode("cp1252")
        self.assertIn(f'ID="{self.project_id}"', project)
        self.assertIn("Document=ThisWorkbook/&H00000000", project)
        self.assertIn("Module=CORE_Small", project)
        directory = decompress(cfb.read("VBA/dir"))
        for module in sample_modules():
            self.assertIn(module.name.encode("cp1252"), directory)
        self.assertEqual(cfb.read("VBA/_VBA_PROJECT")[:2], b"\xcc\x61")

    def test_output_is_deterministic_for_a_project_id(self) -> None:
        first = vba_project.build_vba_project(sample_modules(), self.project_id)
        self.assertEqual(first, vba_project.build_vba_project(sample_modules(), self.project_id))
        self.assertNotEqual(first, vba_project.build_vba_project(sample_modules(), "{" + "1" * 8 + "-0000-0000-0000-" + "0" * 12 + "}"))


@unittest.skipUnless(__import__("importlib").util.find_spec("openpyxl"), "openpyxl not installed")
class DevelopmentWorkbook(unittest.TestCase):
    def test_workbook_embeds_every_component(self) -> None:
        import tempfile
        sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
        import build_dev_workbook
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "dev.xlsm"
            build_dev_workbook.build(target)
            with zipfile.ZipFile(target) as package:
                self.assertIn("xl/vbaProject.bin", package.namelist())
                self.assertIn("macroEnabled", package.read("[Content_Types].xml").decode())
                cfb = CompoundFile(package.read("xl/vbaProject.bin"))
        for path in build_dev_workbook.components():
            with self.subTest(component=path.stem):
                source = decompress(cfb.read(f"VBA/{path.stem}")).decode("cp1252")
                expected = path.read_bytes().decode("cp1252").replace("\r\n", "\n").replace("\n", "\r\n")
                self.assertEqual(source.rstrip("\r\n"), expected.rstrip("\r\n"))


if __name__ == "__main__":
    unittest.main()
