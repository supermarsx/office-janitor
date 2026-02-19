"""!
@brief Automate ``office-janitor`` PyPI releases with ``0.0.x`` patch bumps.
@details The script reads ``src/office_janitor/VERSION``, discovers published
versions from PyPI, picks the next available patch, writes it back to the
version file, builds wheel/sdist artifacts, and optionally uploads them with
Twine. Upload collisions are retried automatically with the next patch.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Iterable, Mapping, Sequence
from urllib import error, request

PATCH_VERSION_RE = re.compile(r"^0\.0\.(\d+)$")


def parse_patch_version(version: str) -> int:
    """!
    @brief Parse a ``0.0.x`` version string and return the patch number.
    @param version Candidate version string.
    @returns Integer patch component.
    @throws ValueError If ``version`` is not in ``0.0.x`` format.
    """

    match = PATCH_VERSION_RE.fullmatch(version.strip())
    if not match:
        raise ValueError(f"Expected 0.0.x version, got: {version!r}")
    return int(match.group(1))


def format_patch_version(patch: int) -> str:
    """!
    @brief Build a ``0.0.x`` version string from an integer patch number.
    @param patch Non-negative patch component.
    @returns Normalized version string.
    @throws ValueError If ``patch`` is negative.
    """

    if patch < 0:
        raise ValueError(f"Patch number must be non-negative, got: {patch}")
    return f"0.0.{patch}"


def read_local_version(version_file: Path) -> str:
    """!
    @brief Read and normalize the local version file content.
    @param version_file Path to ``src/office_janitor/VERSION``.
    @returns Trimmed version string.
    """

    return version_file.read_text(encoding="utf-8").strip()


def write_local_version(version_file: Path, version: str) -> None:
    """!
    @brief Persist the selected version into the local version file.
    @param version_file Path to ``src/office_janitor/VERSION``.
    @param version Version string to store.
    """

    version_file.write_text(f"{version}\n", encoding="utf-8")


def fetch_pypi_release_payload(package_name: str, timeout: float = 10.0) -> Mapping[str, object]:
    """!
    @brief Retrieve package metadata from the PyPI JSON API.
    @param package_name Distribution name used on PyPI.
    @param timeout Network timeout in seconds.
    @returns Parsed JSON payload; empty releases for first-time publish.
    @throws RuntimeError For network/HTTP failures other than 404.
    """

    url = f"https://pypi.org/pypi/{package_name}/json"
    try:
        with request.urlopen(url, timeout=timeout) as response:
            payload = json.load(response)
            if isinstance(payload, Mapping):
                return payload
            raise RuntimeError(f"Unexpected PyPI response type: {type(payload)!r}")
    except error.HTTPError as exc:
        if exc.code == 404:
            return {"releases": {}}
        raise RuntimeError(f"Failed to query PyPI ({exc.code}): {exc.reason}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"Failed to query PyPI: {exc.reason}") from exc


def extract_published_patches(payload: Mapping[str, object]) -> set[int]:
    """!
    @brief Extract all published ``0.0.x`` patch numbers from PyPI metadata.
    @param payload JSON payload from ``/pypi/<name>/json``.
    @returns Set of published patch integers.
    """

    releases = payload.get("releases")
    if not isinstance(releases, Mapping):
        return set()

    patches: set[int] = set()
    for raw_version in releases.keys():
        if not isinstance(raw_version, str):
            continue
        match = PATCH_VERSION_RE.fullmatch(raw_version.strip())
        if match:
            patches.add(int(match.group(1)))
    return patches


def choose_next_patch(local_version: str, published_patches: Iterable[int]) -> int:
    """!
    @brief Compute the next release patch from local and published state.
    @param local_version Current local ``0.0.x`` version.
    @param published_patches Existing ``0.0.x`` patches on PyPI.
    @returns Next patch integer that should be used for release.
    """

    local_patch = parse_patch_version(local_version)
    highest_published = max(published_patches, default=-1)
    return max(local_patch, highest_published) + 1


def expected_artifacts(dist_dir: Path, package_name: str, version: str) -> list[Path]:
    """!
    @brief Build expected wheel and sdist filenames for a release version.
    @param dist_dir Distribution output directory.
    @param package_name PyPI package name (hyphen form).
    @param version Chosen package version.
    @returns Ordered artifact paths (wheel then sdist).
    """

    normalized_name = package_name.replace("-", "_")
    return [
        dist_dir / f"{normalized_name}-{version}-py3-none-any.whl",
        dist_dir / f"{normalized_name}-{version}.tar.gz",
    ]


def build_distributions(python_exe: str, dist_dir: Path) -> None:
    """!
    @brief Build wheel and source distribution into ``dist_dir``.
    @param python_exe Python executable used to run ``python -m build``.
    @param dist_dir Output directory for built artifacts.
    """

    dist_dir.mkdir(parents=True, exist_ok=True)
    command = [python_exe, "-m", "build", "--outdir", str(dist_dir)]
    subprocess.run(command, check=True)


def ensure_release_tools(python_exe: str) -> None:
    """!
    @brief Ensure required packaging tools are installed in current environment.
    @param python_exe Python executable used to run pip.
    """

    command = [python_exe, "-m", "pip", "install", "--upgrade", "build", "twine"]
    subprocess.run(command, check=True)


def upload_distributions(
    python_exe: str,
    artifacts: Sequence[Path],
    repository_url: str | None = None,
) -> None:
    """!
    @brief Upload built artifacts using Twine.
    @param python_exe Python executable used to run ``python -m twine``.
    @param artifacts Built artifact paths to upload.
    @param repository_url Optional custom repository URL (e.g. TestPyPI).
    @throws subprocess.CalledProcessError If upload fails.
    """

    command = [python_exe, "-m", "twine", "upload"]
    if repository_url:
        command.extend(["--repository-url", repository_url])
    command.extend(str(path) for path in artifacts)

    completed = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
    )
    if completed.stdout:
        print(completed.stdout, end="")
    if completed.stderr:
        print(completed.stderr, end="", file=sys.stderr)
    if completed.returncode != 0:
        raise subprocess.CalledProcessError(
            completed.returncode,
            command,
            output=completed.stdout,
            stderr=completed.stderr,
        )


def is_existing_file_upload_error(exc: subprocess.CalledProcessError) -> bool:
    """!
    @brief Determine whether a Twine failure indicates version/file collision.
    @param exc Raised upload exception.
    @returns ``True`` when output suggests the file already exists on registry.
    """

    output = " ".join(part for part in (exc.output, exc.stderr) if part).lower()
    markers = (
        "already exist",
        "already been taken",
        "cannot upload",
        "filename has already been used",
        "this filename has already been used",
    )
    return any(marker in output for marker in markers)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """!
    @brief Parse command-line arguments.
    @param argv Optional custom argv sequence for testing.
    @returns Parsed argument namespace.
    """

    parser = argparse.ArgumentParser(
        description=(
            "Auto-bump 0.0.x version, build artifacts, and optionally upload "
            "office-janitor to PyPI."
        )
    )
    parser.add_argument(
        "--python", default=sys.executable, help="Python executable to run build/twine."
    )
    parser.add_argument("--package-name", default="office-janitor", help="PyPI package name.")
    parser.add_argument(
        "--version-file",
        default="src/office_janitor/VERSION",
        help="Path to VERSION source-of-truth file.",
    )
    parser.add_argument("--dist-dir", default="dist", help="Directory for built artifacts.")
    parser.add_argument(
        "--repository-url",
        default=None,
        help="Custom repository URL (e.g. https://test.pypi.org/legacy/).",
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Upload built artifacts after version bump/build.",
    )
    parser.add_argument(
        "--max-attempts",
        type=int,
        default=5,
        help="Max retries when upload fails because that version already exists.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute and print next version without writing/building/uploading.",
    )
    parser.add_argument(
        "--refresh-tools",
        action="store_true",
        help="Install or update build/twine before releasing.",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    """!
    @brief Execute automatic version bump/build/publish workflow.
    @param argv Optional custom argv sequence.
    @returns Process exit code.
    """

    args = parse_args(argv)
    if args.max_attempts < 1:
        raise ValueError("--max-attempts must be >= 1")

    version_file = Path(args.version_file)
    dist_dir = Path(args.dist_dir)
    current_version = read_local_version(version_file)

    if args.refresh_tools:
        ensure_release_tools(args.python)

    for attempt in range(1, args.max_attempts + 1):
        payload = fetch_pypi_release_payload(args.package_name)
        published_patches = extract_published_patches(payload)
        next_patch = choose_next_patch(current_version, published_patches)
        next_version = format_patch_version(next_patch)

        print(f"[release] selected version: {next_version} (attempt {attempt}/{args.max_attempts})")
        if args.dry_run:
            return 0

        write_local_version(version_file, next_version)
        print(f"[release] updated {version_file} -> {next_version}")

        build_distributions(args.python, dist_dir)
        artifacts = expected_artifacts(dist_dir, args.package_name, next_version)
        missing = [str(path) for path in artifacts if not path.exists()]
        if missing:
            raise RuntimeError("Expected artifacts not found after build: " + ", ".join(missing))

        if not args.publish:
            print("[release] build completed (publish disabled)")
            return 0

        try:
            upload_distributions(args.python, artifacts, repository_url=args.repository_url)
        except subprocess.CalledProcessError as exc:
            if is_existing_file_upload_error(exc) and attempt < args.max_attempts:
                print(
                    "[release] upload version collision detected, retrying with next patch...",
                    file=sys.stderr,
                )
                current_version = next_version
                continue
            raise

        print(f"[release] publish complete: {next_version}")
        return 0

    raise RuntimeError(f"Failed to publish after {args.max_attempts} attempts.")


if __name__ == "__main__":
    raise SystemExit(main())
