#!/usr/bin/env python3
"""Build and verify a hash-locked offline wheelhouse for Windows."""

from __future__ import annotations

import argparse
import json
import os
import platform
import re
import subprocess
import sys
import tempfile
import venv
from email.parser import Parser
from pathlib import Path

try:
    from .lock_dependencies import (
        ARCHITECTURES,
        LOCK_ROOT,
        PYTHON_TAGS,
        REPO_ROOT,
        file_sha256,
        normalize_name,
        wheel_identity,
    )
except ImportError:
    from lock_dependencies import (
        ARCHITECTURES,
        LOCK_ROOT,
        PYTHON_TAGS,
        REPO_ROOT,
        file_sha256,
        normalize_name,
        wheel_identity,
    )


LOCK_NAME_RE = re.compile(r"^windows-cp(311|312|313|314)-(x64|x86|arm64)\.txt$")
REQUIREMENT_RE = re.compile(r"^([A-Za-z0-9_.-]+)==([A-Za-z0-9_.+!-]+)\s*\\$")
HASH_RE = re.compile(r"^\s*--hash=sha256:([0-9a-f]{64})(?:\s*\\)?$")


def current_architecture(machine: str | None = None) -> str:
    if machine is None and os.name == "nt" and sys.maxsize <= 2**32:
        return "x86"
    value = (machine or platform.machine()).strip().lower()
    aliases = {
        "amd64": "x64",
        "x86_64": "x64",
        "x64": "x64",
        "i386": "x86",
        "i686": "x86",
        "x86": "x86",
        "arm64": "arm64",
        "aarch64": "arm64",
    }
    try:
        return aliases[value]
    except KeyError as exc:
        raise ValueError(f"unsupported Windows architecture: {value or 'unknown'}") from exc


def default_lock_path(
    version_info: tuple[int, int] | None = None,
    machine: str | None = None,
) -> Path:
    major, minor = version_info or sys.version_info[:2]
    python_tag = f"{major}{minor}"
    if python_tag not in PYTHON_TAGS:
        raise ValueError("wheelhouse tooling supports CPython 3.11 through 3.14")
    return LOCK_ROOT / f"windows-cp{python_tag}-{current_architecture(machine)}.txt"


def parse_lock_target(path: Path) -> tuple[str, str]:
    match = LOCK_NAME_RE.fullmatch(path.name)
    if not match:
        raise ValueError(
            "lock filename must match windows-cp<311-314>-<x64|x86|arm64>.txt"
        )
    return match.group(1), match.group(2)


def parse_lock(path: Path) -> dict[str, tuple[str, set[str]]]:
    requirements: dict[str, tuple[str, set[str]]] = {}
    current_name: str | None = None
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        requirement = REQUIREMENT_RE.fullmatch(raw_line)
        if requirement:
            display_name, version = requirement.groups()
            current_name = normalize_name(display_name)
            if current_name in requirements:
                raise ValueError(f"{path}: duplicate requirement {display_name}")
            requirements[current_name] = (version, set())
            continue
        digest = HASH_RE.fullmatch(raw_line)
        if digest and current_name:
            requirements[current_name][1].add(digest.group(1))
            continue
        if line in {"--only-binary=:all:", "--require-hashes"}:
            continue
        raise ValueError(f"{path}: unsupported lock line: {raw_line}")
    if not requirements or any(not hashes for _, hashes in requirements.values()):
        raise ValueError(f"{path}: every exact requirement must have a SHA-256 hash")
    return requirements


def verify_wheels(directory: Path, lock_path: Path) -> list[dict[str, object]]:
    expected = parse_lock(lock_path)
    observed: dict[str, dict[str, object]] = {}
    for wheel_path in sorted(directory.glob("*.whl"), key=lambda path: path.name.lower()):
        name, version = wheel_identity(wheel_path)
        normalized = normalize_name(name)
        if normalized in observed:
            raise ValueError(f"duplicate wheel for {name}")
        if normalized not in expected:
            raise ValueError(f"unexpected wheel {wheel_path.name}")
        expected_version, allowed_hashes = expected[normalized]
        digest = file_sha256(wheel_path)
        if version != expected_version:
            raise ValueError(
                f"{wheel_path.name}: version {version} does not match {expected_version}"
            )
        if digest not in allowed_hashes:
            raise ValueError(f"{wheel_path.name}: SHA-256 is not present in the lock")
        observed[normalized] = {
            "name": name,
            "version": version,
            "filename": wheel_path.name,
            "sha256": digest,
            "size": wheel_path.stat().st_size,
        }
    if set(observed) != set(expected):
        missing = sorted(set(expected) - set(observed))
        raise ValueError(f"wheelhouse is missing locked distributions: {missing}")
    return [observed[name] for name in sorted(observed)]


def pip_target_args(python_tag: str, architecture: str) -> list[str]:
    return [
        "--platform",
        ARCHITECTURES[architecture],
        "--python-version",
        python_tag,
        "--implementation",
        "cp",
        "--abi",
        f"cp{python_tag}",
    ]


def install_validation(
    wheelhouse: Path,
    lock_path: Path,
    python_tag: str,
    architecture: str,
) -> None:
    current_tag = f"{sys.version_info.major}{sys.version_info.minor}"
    native = current_tag == python_tag and current_architecture() == architecture
    with tempfile.TemporaryDirectory(prefix="msstorehelper-wheel-test-") as temp_root:
        temp_path = Path(temp_root)
        if native:
            environment = temp_path / "venv"
            venv.EnvBuilder(with_pip=True, clear=True).create(environment)
            python = (
                environment / "Scripts" / "python.exe"
                if os.name == "nt"
                else environment / "bin" / "python"
            )
            subprocess.run(
                [
                    str(python),
                    "-m",
                    "pip",
                    "install",
                    "--disable-pip-version-check",
                    "--no-index",
                    "--find-links",
                    str(wheelhouse),
                    "--require-hashes",
                    "--only-binary=:all:",
                    "--requirement",
                    str(lock_path),
                ],
                cwd=REPO_ROOT,
                check=True,
                timeout=900,
            )
            subprocess.run(
                [str(python), "-m", "unittest", "discover", "-s", "tests"],
                cwd=REPO_ROOT,
                check=True,
                timeout=600,
            )
            return

        target = temp_path / "target"
        subprocess.run(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "--disable-pip-version-check",
                "--no-index",
                "--find-links",
                str(wheelhouse),
                "--require-hashes",
                "--only-binary=:all:",
                "--no-deps",
                "--target",
                str(target),
                *pip_target_args(python_tag, architecture),
                "--requirement",
                str(lock_path),
            ],
            cwd=REPO_ROOT,
            check=True,
            timeout=900,
        )
        installed: dict[str, str] = {}
        for metadata_path in target.glob("*.dist-info/METADATA"):
            metadata = Parser().parsestr(metadata_path.read_text(encoding="utf-8"))
            name = metadata.get("Name")
            version = metadata.get("Version")
            if not name or not version:
                raise ValueError(f"{metadata_path}: missing Name or Version")
            installed[normalize_name(name)] = version
        expected = parse_lock(lock_path)
        if set(installed) != set(expected):
            raise ValueError(
                f"cross-target install metadata mismatch: "
                f"missing={sorted(set(expected) - set(installed))}, "
                f"unexpected={sorted(set(installed) - set(expected))}"
            )
        for name, (expected_version, _) in expected.items():
            if installed[name] != expected_version:
                raise ValueError(
                    f"cross-target {name} version {installed[name]} "
                    f"does not match {expected_version}"
                )


def build_wheelhouse(lock_path: Path, output: Path, *, test: bool) -> Path:
    python_tag, architecture = parse_lock_target(lock_path)
    parse_lock(lock_path)
    if output.exists():
        raise FileExistsError(f"refusing to overwrite existing output: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(
        prefix=f".{output.name}.", dir=output.parent
    ) as temporary:
        staging = Path(temporary)
        subprocess.run(
            [
                sys.executable,
                "-m",
                "pip",
                "download",
                "--disable-pip-version-check",
                "--no-deps",
                "--only-binary=:all:",
                "--require-hashes",
                *pip_target_args(python_tag, architecture),
                "--dest",
                str(staging),
                "--requirement",
                str(lock_path),
            ],
            cwd=REPO_ROOT,
            check=True,
            timeout=900,
        )
        files = verify_wheels(staging, lock_path)
        if test:
            install_validation(staging, lock_path, python_tag, architecture)
        manifest = {
            "schema_version": 1,
            "target": {
                "implementation": "cp",
                "python": f"{python_tag[0]}.{python_tag[1:]}",
                "architecture": architecture,
                "platform": ARCHITECTURES[architecture],
            },
            "lock": lock_path.name,
            "files": files,
        }
        manifest_path = staging / "wheelhouse-manifest.json"
        with manifest_path.open("w", encoding="utf-8", newline="\n") as stream:
            json.dump(manifest, stream, indent=2, sort_keys=True)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        staging.rename(output)
    return output


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--lock",
        type=Path,
        help="target lock; defaults to the running interpreter and architecture",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--test",
        action="store_true",
        help="prove offline installation; native targets also run the full test suite",
    )
    args = parser.parse_args(argv)

    try:
        lock_path = (args.lock or default_lock_path()).resolve(strict=True)
        output = args.output.resolve()
        built = build_wheelhouse(lock_path, output, test=args.test)
    except (OSError, ValueError, subprocess.CalledProcessError) as exc:
        print(f"wheelhouse build failed: {exc}", file=sys.stderr)
        return 1
    print(f"verified wheelhouse: {built}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
