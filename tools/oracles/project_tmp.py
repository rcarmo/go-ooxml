"""Project-owned disposable LibreOffice profile roots for standalone/CI oracles."""
import os
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
        if os.stat(path).st_uid != os.geteuid():
            return False
    parent = path
    while not os.path.exists(parent):
        parent = os.path.dirname(parent)
    return os.path.isdir(parent) and os.access(parent, os.W_OK | os.X_OK)


def _ci():
    if os.environ.get("CI", "").lower() not in ("", "0", "false"):
        return True
    return any(os.environ.get(name, "").lower() == "true" for name in
               ("GITHUB_ACTIONS", "GITLAB_CI", "TF_BUILD", "CIRCLECI"))


def project_root():
    base = os.environ.get("PROJECT_TMP_BASE")
    explicit = os.environ.get("PROJECT_TMP_ROOT")
    base_root = None
    if base is not None:
        if not base:
            raise ValueError("PROJECT_TMP_BASE must not be empty")
        base_root = os.path.join(base.rstrip("/"), PROJECT)
        if not _usable(base_root):
            raise ValueError("Invalid PROJECT_TMP_BASE (absolute usable base required)")
    if explicit is not None:
        candidate = explicit.rstrip("/")
        if os.path.basename(candidate) != PROJECT or not _usable(candidate):
            raise ValueError("Invalid PROJECT_TMP_ROOT (absolute usable project-named root required)")
        if base_root is not None and candidate != base_root:
            raise ValueError("Conflicting PROJECT_TMP_BASE and PROJECT_TMP_ROOT")
        return candidate
    if base_root is not None:
        return base_root
    original_tmp = os.environ.get("PROJECT_ORIGINAL_TMPDIR", os.environ.get("TMPDIR", ""))
    bases = (os.environ.get("RUNNER_TEMP"), original_tmp, "/tmp") if _ci() else ("/workspace/tmp", "/tmp")
    for candidate_base in bases:
        if candidate_base and _usable(os.path.join(candidate_base.rstrip("/"), PROJECT)):
            return os.path.join(candidate_base.rstrip("/"), PROJECT)
    raise ValueError("No usable project-owned temporary root")


def oracle_profile():
    root = project_root()
    scratch = os.environ.get("OOXML_ORACLE_SCRATCH") or os.path.join(root, "runs", "oracles", uuid.uuid4().hex)
    if not _usable(scratch) or os.path.commonpath((root, scratch)) != root or os.path.relpath(scratch, root).split(os.sep)[0] != "runs":
        raise ValueError("Oracle scratch must be an isolated project-owned runs/ directory")
    os.makedirs(scratch, exist_ok=True)
    return os.path.join(scratch, "profile-" + uuid.uuid4().hex)
