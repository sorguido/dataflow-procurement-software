"""Helpers for restart command resolution and post-mainloop spawn."""

from __future__ import annotations

import os
import subprocess
import sys


def resolve_restart_script_path(*, file_value, argv0, cwd, executable, is_frozen):
    """Resolve the best script/executable path for restart."""
    script_path = None

    try:
        if file_value:
            candidate = os.path.abspath(file_value)
            if os.path.exists(candidate) and candidate.endswith(".py"):
                script_path = candidate
    except (NameError, AttributeError):
        pass

    if not script_path or not os.path.exists(script_path):
        if argv0:
            if os.path.exists(argv0):
                script_path = os.path.abspath(argv0)
            else:
                possible_path = os.path.join(cwd, argv0)
                if os.path.exists(possible_path):
                    script_path = os.path.abspath(possible_path)
                elif is_frozen:
                    script_path = executable
                else:
                    script_path = argv0

    if not script_path or (not os.path.exists(script_path) and not is_frozen):
        script_path = executable

    return script_path


def build_restart_command(*, python_executable, script_path, is_frozen):
    """Build restart command preserving current script/executable rules."""
    if script_path.endswith(".py") or (not is_frozen and script_path != python_executable):
        return [python_executable, script_path]
    return [script_path]


def launch_post_mainloop_restart(*, cmd, cwd, platform_name):
    """Launch detached restart process after Tk mainloop exits."""
    if platform_name == "win32":
        subprocess.Popen(
            cmd,
            cwd=cwd,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
            close_fds=True,
        )
    else:
        subprocess.Popen(cmd, cwd=cwd, start_new_session=True)
