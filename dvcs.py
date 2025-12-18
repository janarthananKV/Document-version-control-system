#!/usr/bin/env python3

from __future__ import annotations
import argparse
import dataclasses
import datetime as dt
import io
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple
import zipfile
import xml.etree.ElementTree as ET
import difflib
from io import BytesIO


# Utilities
def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def repo_dir_for(document_path: str | Path) -> Path:
    p = Path(document_path)
    base = f".{p.name}.repo"  # keep per-document repo next to file
    return p.with_name(base)


def ensure_dir(d: Path) -> None:
    d.mkdir(parents=True, exist_ok=True)


def which(cmd: str) -> Optional[str]:
    return shutil.which(cmd)


# DOCX helpers for human-readable diffs
_DOCX_MAIN = "word/document.xml"
_VOLATILE_ATTRS = re.compile(r"\s+w:rsid\w+=\"[^\"]+\"")


def extract_docx_xml(docx_bytes: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as z:
        xml_bytes = z.read(_DOCX_MAIN)
    return xml_bytes.decode("utf-8", errors="replace")


def normalize_docx_xml(xml: str) -> str:
    # Strip volatile revision/session ids and normalize whitespace/newlines
    xml = _VOLATILE_ATTRS.sub("", xml)
    xml = xml.replace("\r\n", "\n").replace("\r", "\n")
    # Optional: collapse multiple spaces between tags
    xml = re.sub(r">\s+<", "><", xml)
    return xml


# Storage metadata
@dataclasses.dataclass
class VersionEntry:
    version: int
    kind: str  # "snapshot" | "delta"
    file: str  # file name in repo dir
    message: str
    created_at: str
    base_version: Optional[int] = None  # for deltas, the version it applies from


@dataclasses.dataclass
class RepoState:
    document_type: str  # "docx" | "pdf"
    snapshot_interval: int = 5
    versions: List[VersionEntry] = dataclasses.field(default_factory=list)

    @staticmethod
    def load(path: Path) -> Optional["RepoState"]:
        meta = path / "metadata.json"
        if not meta.exists():
            return None
        data = json.loads(meta.read_text("utf-8"))
        versions = [VersionEntry(**v) for v in data["versions"]]
        return RepoState(document_type=data["document_type"],
                         snapshot_interval=data.get("snapshot_interval", 5),
                         versions=versions)

    def save(self, path: Path) -> None:
        data = {
            "document_type": self.document_type,
            "snapshot_interval": self.snapshot_interval,
            "versions": [dataclasses.asdict(v) for v in self.versions],
        }
        (path / "metadata.json").write_text(json.dumps(data, indent=2), "utf-8")


# Core DVCS logic
class DVCS:
    def __init__(self, file_path: str, doc_type: str, snapshot_interval: int = 5):
        self.file_path = Path(file_path)
        self.doc_type = doc_type
        self.repo_path = repo_dir_for(self.file_path)
        self.snapshot_interval = snapshot_interval
        ensure_dir(self.repo_path)
        self.state = RepoState.load(self.repo_path)

    # Init 
    def init(self, message: str) -> None:
        if self.state is not None:
            print("Repository already initialized.")
            return
        data = self.file_path.read_bytes()
        vfile = self._version_file(1, snapshot=True)
        vfile.write_bytes(data)
        entry = VersionEntry(version=1, kind="snapshot", file=vfile.name,
                             message=message, created_at=now_iso())
        self.state = RepoState(document_type=self.doc_type,
                               snapshot_interval=self.snapshot_interval,
                               versions=[entry])
        self.state.save(self.repo_path)
        print(f"Initialized repo at {self.repo_path}")

    # Add version
    def add(self, message: str) -> None:
        self._require_initialized()
        data = self.file_path.read_bytes()
        next_ver = self.state.versions[-1].version + 1
        # Decide snapshot vs delta
        if (next_ver - 1) % self.state.snapshot_interval == 0:
            # make a snapshot
            vfile = self._version_file(next_ver, snapshot=True)
            vfile.write_bytes(data)
            entry = VersionEntry(version=next_ver, kind="snapshot", file=vfile.name,
                                 message=message, created_at=now_iso())
            self.state.versions.append(entry)
            self.state.save(self.repo_path)
            print(f"Added snapshot v{next_ver}")
            return

        # else make a delta from nearest base snapshot
        base_idx, base_entry = self._nearest_snapshot_before(next_ver)
        base_bytes = (self.repo_path / base_entry.file).read_bytes()

        # Reconstruct prior version (next_ver-1) to use as delta source
        prior_bytes = self._reconstruct_bytes(next_ver - 1)

        # try xdelta3
        # try xdelta3
        xd = which("xdelta3")
        if xd:
            delta_path = self._version_file(next_ver, snapshot=False)
            # Create a temp file for prior version to feed to xdelta3
            tmp_prior = self.repo_path / ".prior.tmp"
            tmp_prior.write_bytes(prior_bytes)
            tmp_new = self.repo_path / ".new.tmp"
            tmp_new.write_bytes(data)
            try:
                # xdelta3 -e -s PRIOR NEW DELTA
                subprocess.run([xd, "-e", "-s", str(tmp_prior), str(tmp_new), str(delta_path)],
                            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

                # prepare fields properly
                entry = VersionEntry(
                    version=next_ver,
                    kind="delta",
                    file=delta_path.name,
                    message=message,
                    created_at=now_iso(),
                    base_version=next_ver - 1   # <-- just pass it directly
                )
                self.state.versions.append(entry)
                self.state.save(self.repo_path)
                print(f"Added delta v{next_ver} (xdelta3)")
                return
            except subprocess.CalledProcessError as e:
                print("xdelta3 failed, storing snapshot instead:", e)
            finally:
                for p in (tmp_prior, tmp_new):
                    if p.exists():
                        p.unlink()


        # fallback: store full snapshot
        vfile = self._version_file(next_ver, snapshot=True)
        vfile.write_bytes(data)
        entry = VersionEntry(version=next_ver, kind="snapshot", file=vfile.name,
                             message=message + " (fallback snapshot)", created_at=now_iso())
        self.state.versions.append(entry)
        self.state.save(self.repo_path)
        print(f"Added snapshot v{next_ver} (fallback)")

    # Get / Reconstruct
    def get(self, version: int, output: str) -> None:
        self._require_initialized()
        out_path = Path(output)
        bytes_ = self._reconstruct_bytes(version)
        out_path.write_bytes(bytes_)
        print(f"Wrote v{version} -> {out_path}")

    # Revert file in-place
    def revert(self, version: int) -> None:
        self._require_initialized()
        bytes_ = self._reconstruct_bytes(version)
        self.file_path.write_bytes(bytes_)
        print(f"Reverted {self.file_path.name} to v{version}")

    # History
    def history(self) -> None:
        self._require_initialized()
        print(f"History for {self.file_path.name} (type={self.state.document_type}):")
        for v in self.state.versions:
            base = f" base={v.base_version}" if v.base_version else ""
            print(f"  v{v.version:04d} [{v.kind}]{base}  {v.created_at}  {v.message}")

    # DOCX human-readable diff
    def show_diff(self, v1: int, v2: int) -> None:
        self._require_initialized()
        if self.state.document_type != "docx":
            print("Human-readable diff is only supported for DOCX.")
            return

        def extract_text_from_docx_bytes(docx_bytes):
            texts = []
            with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
                with z.open("word/document.xml") as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                    for t in root.findall(".//w:t", ns):
                        texts.append(t.text or "")
            return texts

        # reconstruct both versions
        b1 = self._reconstruct_bytes(v1)
        b2 = self._reconstruct_bytes(v2)

        # extract human-readable text
        x1 = extract_text_from_docx_bytes(b1)
        x2 = extract_text_from_docx_bytes(b2)

        # run line-based diff
        diff = difflib.unified_diff(
            x1, x2, fromfile=f"v{v1}", tofile=f"v{v2}", lineterm=""
        )
        text = "\n".join(diff)

        if not text.strip():
            print("No significant textual changes detected.")
        else:
            print(text)


    # Internals
    def _reconstruct_bytes(self, target_version: int) -> bytes:
        if target_version < 1 or target_version > self.state.versions[-1].version:
            raise ValueError(f"version must be in [1, {self.state.versions[-1].version}]")
        # Find nearest snapshot at or before target
        snap_idx, snap_entry = self._nearest_snapshot_at_or_before(target_version)
        data = (self.repo_path / snap_entry.file).read_bytes()
        # Apply deltas from snapshot+1 .. target
        for v in range(snap_entry.version + 1, target_version + 1):
            entry = self._entry_for(v)
            if entry.kind == "snapshot":
                data = (self.repo_path / entry.file).read_bytes()
                continue
            xd = which("xdelta3")
            if xd:
                # Write current data to temp, apply delta
                tmp_base = self.repo_path / ".base.tmp"
                tmp_base.write_bytes(data)
                tmp_out = self.repo_path / ".out.tmp"
                try:
                    subprocess.run([xd, "-d", "-s", str(tmp_base),
                                    str(self.repo_path / entry.file), str(tmp_out)],
                                   check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    data = tmp_out.read_bytes()
                except subprocess.CalledProcessError as e:
                    raise RuntimeError(f"Failed to apply delta for v{v}: {e}")
                finally:
                    for p in (tmp_base, tmp_out):
                        if p.exists():
                            p.unlink()
            else:
                # No xdelta3 available; deltas shouldn't exist, but if they do, fail fast
                raise RuntimeError("xdelta3 not available to apply delta; cannot reconstruct.")
        return data

    def _nearest_snapshot_before(self, next_ver: int) -> Tuple[int, VersionEntry]:
        # Find nearest snapshot strictly before next_ver
        for i in range(len(self.state.versions) - 1, -1, -1):
            v = self.state.versions[i]
            if v.version < next_ver and v.kind == "snapshot":
                return i, v
        # Should never happen because v1 is a snapshot
        raise RuntimeError("No snapshot found before requested version.")

    def _nearest_snapshot_at_or_before(self, ver: int) -> Tuple[int, VersionEntry]:
        for i in range(len(self.state.versions) - 1, -1, -1):
            v = self.state.versions[i]
            if v.version <= ver and v.kind == "snapshot":
                return i, v
        raise RuntimeError("No snapshot found at or before requested version.")

    def _entry_for(self, ver: int) -> VersionEntry:
        for v in self.state.versions:
            if v.version == ver:
                return v
        raise KeyError(f"No entry for version {ver}")

    def _version_file(self, ver: int, snapshot: bool) -> Path:
        suffix = ".snapshot" if snapshot else ".delta"
        name = f"v{ver:04d}{suffix}"
        return self.repo_path / name

    def _require_initialized(self) -> None:
        if self.state is None:
            raise RuntimeError("Repository not initialized. Run `init` first.")


# CLI
def main():
    parser = argparse.ArgumentParser(description="Hybrid DVCS for DOCX/PDF (snapshots + binary deltas + DOCX diff viewer)")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_init = sub.add_parser("init", help="Initialize repository for a file")
    p_init.add_argument("file", help="Path to DOCX/PDF")
    p_init.add_argument("-t", "--type", choices=["docx", "pdf"], required=True)
    p_init.add_argument("-m", "--message", default="Initial version")
    p_init.add_argument("--interval", type=int, default=5, help="Snapshot interval (default: 5)")

    p_add = sub.add_parser("add", help="Add current file contents as a new version")
    p_add.add_argument("file")
    p_add.add_argument("-t", "--type", choices=["docx", "pdf"], required=True)
    p_add.add_argument("-m", "--message", default="Update")

    p_hist = sub.add_parser("history", help="Show history for a file")
    p_hist.add_argument("file")

    p_get = sub.add_parser("get", help="Reconstruct a version to an output path")
    p_get.add_argument("file")
    p_get.add_argument("version", type=int)
    p_get.add_argument("output")

    p_rev = sub.add_parser("revert", help="Revert the working file to a specific version")
    p_rev.add_argument("file")
    p_rev.add_argument("version", type=int)

    p_diff = sub.add_parser("show-diff", help="Human-readable diff between two DOCX versions (normalized XML)")
    p_diff.add_argument("file")
    p_diff.add_argument("v1", type=int)
    p_diff.add_argument("v2", type=int)

    args = parser.parse_args()

    if args.cmd == "init":
        dvcs = DVCS(args.file, args.type, snapshot_interval=args.interval)
        dvcs.init(message=args.message)
    elif args.cmd == "add":
        dvcs = DVCS(args.file, args.type)
        dvcs.add(message=args.message)
    elif args.cmd == "history":
        # doc type is saved in metadata; infer it
        dvcs = _open_inferred(args.file)
        dvcs.history()
    elif args.cmd == "get":
        dvcs = _open_inferred(args.file)
        dvcs.get(args.version, args.output)
    elif args.cmd == "revert":
        dvcs = _open_inferred(args.file)
        dvcs.revert(args.version)
    elif args.cmd == "show-diff":
        dvcs = _open_inferred(args.file)
        dvcs.show_diff(args.v1, args.v2)
    else:
        parser.error("Unknown command")


def _open_inferred(file_path: str) -> DVCS:
    # Open repo and read metadata to infer type
    rdir = repo_dir_for(file_path)
    state = RepoState.load(rdir)
    if state is None:
        raise RuntimeError("Repository not initialized for this file.")
    return DVCS(file_path, state.document_type, snapshot_interval=state.snapshot_interval)


if __name__ == "__main__":
    main()

