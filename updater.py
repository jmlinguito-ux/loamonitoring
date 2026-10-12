"""In-app updates from GitHub Releases (Velopack).

Same design as the Ground Data Processing Tool updater: an installed build checks
the latest GitHub release in the background, the user approves the download, and
"Update and restart" lets Velopack swap the files and reopen the app.

Builds Velopack does not manage (running from source, a portable exe, or the
PyInstaller dist folder) report the "disabled" status instead of an error; those
copies are updated by installing the release's Setup.exe.
"""

from __future__ import annotations

import threading
from typing import Callable, Optional

# Static Velopack feed: GitHub serves the newest release's assets (including
# releases.win.json) under /releases/latest/download/. This avoids GitHub's
# anonymous REST API limit (60 requests/hour per IP) and needs no token, but the
# repository must be public.
UPDATE_FEED_URL = "https://github.com/jmlinguito-ux/loamonitoring/releases/latest/download/"


def run_startup_hooks() -> None:
    """Let Velopack handle its install/update/uninstall command line, then return.

    Must run before the window is created: during install and update Velopack
    starts the exe with special arguments and expects it to exit quickly.
    """
    try:
        import velopack
    except ImportError:
        return  # Development from source does not need the updater SDK.
    velopack.App().run()


def _versions_equal(a: str, b: str) -> bool:
    return bool(a) and a.strip().lstrip("vV") == (b or "").strip().lstrip("vV")


class UpdateService:
    """Background update check / download / apply, reported through ``emit``.

    State fields: status (idle, checking, up_to_date, available, downloading,
    ready, installing, error, disabled), active (Velopack manages this install),
    currentVersion, latestVersion, releaseNotes, progress (0-100), message, and
    quiet (an automatic check whose failure should not raise a toast).
    """

    def __init__(self, current_version: str, emit: Callable[[dict], None],
                 feed_url: Optional[str] = None):
        self._emit = emit
        self._lock = threading.Lock()
        self._manager = None
        self._update_info = None
        self._busy = False
        self._feed_url = (UPDATE_FEED_URL if feed_url is None else feed_url).strip()
        configured = bool(self._feed_url)
        self._state = {
            "status": "idle",
            "configured": configured,
            "active": False,
            "currentVersion": str(current_version or ""),
            "latestVersion": "",
            "releaseNotes": "",
            "progress": 0,
            "message": "" if configured else "Update feed is not configured in this build.",
            "quiet": True,
        }
        if configured:
            self._initialise_manager()

    def _initialise_manager(self) -> None:
        try:
            import velopack
        except ImportError:
            self._state["message"] = "Running from source: in-app updates are off."
            return
        try:
            # HttpSource explicitly: a bare github.com URL is auto-detected as
            # the GitHub REST source, which rewrites this static feed path into
            # an invalid API URL (404).
            manager = velopack.UpdateManager(velopack.HttpSource(self._feed_url))
            current = manager.get_current_version()
            if current:
                self._state["currentVersion"] = str(current)
            self._manager = manager
            self._state["active"] = True
        except Exception:  # Velopack raises here for portable / unmanaged builds.
            self._state["message"] = (
                "This copy was not installed with Setup.exe, so in-app updates are off."
            )

    # -- state -------------------------------------------------------------
    def state(self) -> dict:
        with self._lock:
            return dict(self._state)

    def _set(self, **changes) -> dict:
        with self._lock:
            self._state.update(changes)
            snapshot = dict(self._state)
        try:
            self._emit(snapshot)
        except Exception:
            pass
        return snapshot

    def _claim(self) -> bool:
        """Reserve the single worker slot; False if a check/download is running."""
        with self._lock:
            if self._busy:
                return False
            self._busy = True
            return True

    def _spawn(self, target, name: str) -> None:
        def run():
            try:
                target()
            finally:
                with self._lock:
                    self._busy = False

        threading.Thread(target=run, name=name, daemon=True).start()

    # -- actions -----------------------------------------------------------
    def check(self, manual: bool = True) -> dict:
        if self._manager is None:
            return self._set(status="disabled", quiet=not manual,
                             message=self._state["message"] or "Updates are unavailable in this build.")
        if self._state["status"] in ("ready", "installing"):
            return self.state()  # An update is already downloaded; keep it.

        def worker():
            try:
                info = self._manager.check_for_updates()
                asset = getattr(info, "TargetFullRelease", None) if info else None
                latest = str(getattr(asset, "Version", "") or "") if asset else ""
                if not asset or _versions_equal(latest, self._state["currentVersion"]):
                    self._update_info = None
                    changes = dict(status="up_to_date", latestVersion="", releaseNotes="",
                                   message="You are running the latest version.")
                else:
                    self._update_info = info
                    changes = dict(status="available", latestVersion=latest,
                                   releaseNotes=str(getattr(asset, "NotesMarkdown", "") or ""),
                                   message=f"Version {latest} is available.")
            except Exception as exc:
                changes = dict(status="error", message=f"Could not check for updates: {exc}")
            self._set(quiet=not manual, **changes)

        if not self._claim():
            return self.state()
        self._set(status="checking", message="Checking for updates…", progress=0, quiet=not manual)
        self._spawn(worker, "update-check")
        return self.state()

    def download(self) -> dict:
        if self._manager is None or self._update_info is None:
            return self._set(status="error", quiet=False, message="Check for an update before downloading.")

        def progress(value):
            try:
                percent = max(0, min(100, int(value)))
            except (TypeError, ValueError):
                percent = 0
            self._set(status="downloading", progress=percent, message=f"Downloading update… {percent}%")

        def worker():
            try:
                self._manager.download_updates(self._update_info, progress)
                changes = dict(status="ready", progress=100,
                               message="Update downloaded. Restart to finish installing it.")
            except Exception as exc:
                changes = dict(status="error", message=f"Could not download the update: {exc}")
            self._set(quiet=False, **changes)

        if not self._claim():
            return self.state()
        self._set(status="downloading", message="Downloading update…", progress=0, quiet=False)
        self._spawn(worker, "update-download")
        return self.state()

    def apply_and_restart(self) -> dict:
        if self._manager is None or self._update_info is None or self._state.get("status") != "ready":
            return self._set(status="error", quiet=False, message="Download the update before restarting.")
        self._set(status="installing", message="Closing the app and installing the update…")
        try:
            # The SDK owns the whole lifecycle: it starts the updater, which waits
            # for this process to exit, swaps the files and relaunches the app.
            self._manager.apply_updates_and_restart(self._update_info)
            return self.state()
        except Exception as exc:
            return self._set(status="error", quiet=False, message=f"Could not start the installer: {exc}")
