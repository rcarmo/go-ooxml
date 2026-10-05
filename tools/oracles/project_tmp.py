"""Project-owned disposable LibreOffice profile roots for standalone/CI oracles."""
import os
import tempfile
import uuid

PROJECT = "go-ooxml"


def _usable(path):
    if not path or not os.path.isabs(path) or os.path.normpath(path) != path:
        return False
    # /workspace can be a host-provided mount alias; reject other symlink ancestry.
    part = path
    while part != "/":
        if part != "/workspace" and os.path.islink(part):
            return False
        part = os.path.dirname(part)
    if os.path.exists(path):
        if not os.path.isdir(path) or not os.access(path, os.W_OK | os.X_OK):
            return False
        try:
            if os.stat(path).st_uid != os.geteuid():
                return False
        except OSError:
            return False
    parent = path
    while not os.path.exists(parent):
        parent = os.path.dirname(parent)
    return os.path.isdir(parent) and os.access(parent, os.W_OK | os.X_OK)


def project_root():
    explicit = os.environ.get("PROJECT_TMP_ROOT")
    if explicit is not None:
        if os.path.basename(explicit.rstrip("/")) != PROJECT or not _usable(explicit.rstrip("/")):
            raise ValueError("Invalid PROJECT_TMP_ROOT (absolute, usable, project-named, non-symlink required)")
        return explicit.rstrip("/")
    # Resolve platform base before mutating TMPDIR in the child process.
    original_tmp = os.environ.get("TMPDIR")
    # tempfile.gettempdir() caches its first lookup; use the platform base
    # only after considering the unchanged incoming TMPDIR and runner path.
    for base in ("/workspace/tmp", os.environ.get("RUNNER_TEMP"), original_tmp, tempfile.gettempdir()):
        if base and os.path.isabs(base) and _usable(os.path.join(base, PROJECT)):
            return os.path.join(base, PROJECT)
    raise ValueError("No usable project-owned temporary root")


def oracle_profile():
    root = project_root()
    scratch = os.environ.get("OOXML_ORACLE_SCRATCH") or os.path.join(root, "runs", "oracles", uuid.uuid4().hex)
    if not _usable(scratch) or os.path.commonpath((root, scratch)) != root or os.path.relpath(scratch, root).split(os.sep)[0] != "runs":
        raise ValueError("Oracle scratch must be an isolated project-owned runs/ directory")
    os.makedirs(scratch, exist_ok=True)
    return os.path.join(scratch, "profile-" + uuid.uuid4().hex)
