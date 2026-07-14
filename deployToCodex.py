#!/usr/bin/env python3
"""Publish the OneNote MCP server and register it with Codex."""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
PROJECT = REPO_ROOT / "src" / "OneNoteMcp" / "OneNoteMcp.csproj"


def find_tool(name: str) -> str:
    resolved = shutil.which(name)
    if not resolved:
        sys.exit(f"ERROR: '{name}' not found on PATH.")
    return resolved


def run(cmd: list[str], *, check: bool, quiet: bool = False) -> int:
    printable = " ".join(cmd)
    if not quiet:
        print(f"    $ {printable}")
    result = subprocess.run(cmd)
    if check and result.returncode != 0:
        sys.exit(f"ERROR: command failed (exit {result.returncode}): {printable}")
    return result.returncode


def main() -> None:
    default_install = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "OneNoteMcp"
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--install-dir", type=Path, default=default_install,
                        help="Publish directory (default: %%LOCALAPPDATA%%\\OneNoteMcp).")
    parser.add_argument("--configuration", default="Release")
    parser.add_argument("--server-name", default="onenote")
    args = parser.parse_args()

    if not PROJECT.exists():
        sys.exit(f"ERROR: project not found: {PROJECT}")
    install_dir = args.install_dir.resolve()
    exe_path = install_dir / "OneNoteMcp.exe"
    codex = find_tool("codex")
    dotnet = find_tool("dotnet")

    print("==> Deploy OneNote MCP to Codex")
    print(f"    Project    : {PROJECT}")
    print(f"    InstallDir : {install_dir}")
    print(f"    Config     : {args.configuration}")
    print(f"    ServerName : {args.server_name}")
    run([codex, "mcp", "remove", args.server_name], check=False, quiet=True)
    if os.name == "nt":
        run(["taskkill", "/F", "/IM", "OneNoteMcp.exe"], check=False, quiet=True)
    run([dotnet, "publish", str(PROJECT), "-c", args.configuration, "-o", str(install_dir)], check=True)
    if not exe_path.exists():
        sys.exit(f"ERROR: expected {exe_path} after publish; not found.")
    run([codex, "mcp", "add", args.server_name, "--", str(exe_path)], check=True)
    run([codex, "mcp", "get", args.server_name], check=False, quiet=True)
    print("Restart Codex sessions to pick up the deployed server.")


if __name__ == "__main__":
    main()
