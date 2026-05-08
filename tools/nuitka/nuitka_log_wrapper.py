"""Stream Nuitka stdout/stderr to the console and a UTF-8 log file; exit with Nuitka's code."""
from __future__ import annotations

import subprocess
import sys
from datetime import datetime


def main() -> int:
    if len(sys.argv) < 3:
        print(
            "usage: nuitka_log_wrapper.py <logfile> ...nuitka CLI args...",
            file=sys.stderr,
        )
        return 2
    log_path = sys.argv[1]
    nuitka_argv = sys.argv[2:]
    exe = sys.executable
    header = (
        f"========== {datetime.now().isoformat(sep=' ', timespec='seconds')} "
        f"nuitka (via {exe}) ==========\n"
    )
    with open(log_path, "w", encoding="utf-8", errors="replace", newline="\n") as logf:
        logf.write(header)
        logf.flush()
        sys.stdout.write(header)
        sys.stdout.flush()
        with subprocess.Popen(
            [exe, "-m", "nuitka", *nuitka_argv],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            errors="replace",
            bufsize=1,
        ) as p:
            assert p.stdout is not None
            try:
                for line in p.stdout:
                    sys.stdout.write(line)
                    sys.stdout.flush()
                    logf.write(line)
                    logf.flush()
            finally:
                ret = p.wait()
            return ret


if __name__ == "__main__":
    raise SystemExit(main())
