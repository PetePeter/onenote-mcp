#!/usr/bin/env python3
"""Publish the OneNote MCP server and register it with Claude Code at USER scope.

Solves the "can't rebuild the server while it's in use" file-lock: Claude runs a
*published* copy in a stable directory, decoupled from the dev tree's bin\\Debug.
You can then rebuild / test src\\ freely while the registered server keeps running
the last-deployed binary -- it only changes when you re-run this script.

Steps: unregister -> stop any running server (release exe lock) -> dotnet publish
-> register the published exe at user scope -> verify.

Usage:
    python deployToClaude.py
    python deployToClaude.py --install-dir "C:\\Tools\\onenote-mcp"
    python deployToClaude.py --configuration Debug --server-name onenote
"""
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
    """Resolve a CLI (claude/dotnet) to a full path, honoring Windows PATHEXT."""
    resolved = shutil.which(name)
    if not resolved:
        sys.exit(f"ERROR: '{name}' not found on PATH.")
    return resolved


def run(cmd: list[str], *, check: bool, quiet: bool = False) -> int:
    """Run a command, streaming output. When check, a non-zero exit aborts."""
    printable = " ".join(cmd)
    if not quiet:
        print(f"    $ {printable}")
    result = subprocess.run(cmd)
    if check and result.returncode != 0:
        sys.exit(f"ERROR: command failed (exit {result.returncode}): {printable}")
    return result.returncode


def main() -> None:
    default_install = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "OneNoteMcp"

    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--install-dir", type=Path, default=default_install,
                        help="Where to publish the standalone server "
                             "(default: %%LOCALAPPDATA%%\\OneNoteMcp).")
    parser.add_argument("--configuration", default="Release", help="Build configuration (default: Release).")
    parser.add_argument("--server-name", default="onenote", help="MCP server name to register (default: onenote).")
    args = parser.parse_args()

    if not PROJECT.exists():
        sys.exit(f"ERROR: project not found: {PROJECT}")

    install_dir: Path = args.install_dir.resolve()
    exe_path = install_dir / "OneNoteMcp.exe"
    claude = find_tool("claude")
    dotnet = find_tool("dotnet")

    print("==> Deploy OneNote MCP to Claude Code (user scope)")
    print(f"    Project    : {PROJECT}")
    print(f"    InstallDir : {install_dir}")
    print(f"    Config     : {args.configuration}")
    print(f"    ServerName : {args.server_name}")

    # 1. Unregister first so Claude Code stops spawning the old server (releases
    #    the exe lock). A missing registration is harmless -> check=False.
    print(f"==> Removing existing '{args.server_name}' registration (user + any local scope)...")
    run([claude, "mcp", "remove", args.server_name, "--scope", "user"], check=False, quiet=True)
    run([claude, "mcp", "remove", args.server_name], check=False, quiet=True)

    # 2. Stop any running server so publish can overwrite its files. taskkill by
    #    image name is dependency-free; deploy wants the server stopped anyway.
    if os.name == "nt":
        print("==> Stopping any running OneNoteMcp.exe (release file lock)...")
        run(["taskkill", "/F", "/IM", "OneNoteMcp.exe"], check=False, quiet=True)

    # 3. Publish a framework-dependent build to the stable install directory.
    print("==> Publishing...")
    run([dotnet, "publish", str(PROJECT), "-c", args.configuration, "-o", str(install_dir)], check=True)
    if not exe_path.exists():
        sys.exit(f"ERROR: expected {exe_path} after publish; not found.")

    # 4. Register the published exe at user scope (available in every project).
    print(f"==> Registering '{args.server_name}' at user scope -> {exe_path}")
    run([claude, "mcp", "add", args.server_name, "--scope", "user", "--", str(exe_path)], check=True)

    # 5. Verify.
    print("==> Done. Verifying:")
    run([claude, "mcp", "get", args.server_name], check=False, quiet=True)
    print("\nRestart your Claude Code sessions to pick up the deployed server.")


if __name__ == "__main__":
    main()
