#!/usr/bin/env python3
import argparse
import difflib
import hashlib
import os
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_ID_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_TYPE_NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
REL_TYPE_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
REL_TYPE_COMMENTS = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
REL_TYPE_AUTHORS = "http://schemas.microsoft.com/office/2018/10/relationships/authors"


def read_text(zf, name):
    data = zf.read(name)
    return data.decode("utf-8", errors="replace")


def strip_ns(tag):
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def parse_relationships(data):
    rels = []
    root = ET.fromstring(data.lstrip(b"\xef\xbb\xbf"))
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rels.append({
            "id": rel.attrib.get("Id", ""),
            "type": rel.attrib.get("Type", ""),
            "target": rel.attrib.get("Target", ""),
            "target_mode": rel.attrib.get("TargetMode", ""),
        })
    return rels


def parse_content_types(data):
    overrides = {}
    root = ET.fromstring(data.lstrip(b"\xef\xbb\xbf"))
    for child in root:
        if strip_ns(child.tag) != "Override":
            continue
        part = child.attrib.get("PartName", "")
        overrides[part] = child.attrib.get("ContentType", "")
    return overrides


def source_from_rels_path(rels_path):
    if rels_path == "_rels/.rels":
        return ""
    rels_dir = posixpath.dirname(rels_path)
    if not rels_dir.endswith("/_rels"):
        return ""
    base = posixpath.basename(rels_path)
    if not base.endswith(".rels"):
        return ""
    source_dir = rels_dir[: -len("/_rels")]
    source_name = base[: -len(".rels")]
    if source_dir:
        return posixpath.join(source_dir, source_name)
    return source_name


def resolve_target(source, target):
    if target.startswith("/"):
        return target.lstrip("/")
    source_dir = posixpath.dirname(source)
    if source_dir in ("", "."):
        return target
    return posixpath.normpath(posixpath.join(source_dir, target))


def load_relationships(zf):
    rels_by_source = {}
    for name in zf.namelist():
        if not name.endswith(".rels"):
            continue
        source = source_from_rels_path(name)
        if source == "":
            source = ""
        rels_by_source[source] = parse_relationships(zf.read(name))
    return rels_by_source


def diff_zips(path_a, path_b, show_diff=False, diff_limit=200):
    with zipfile.ZipFile(path_a) as za, zipfile.ZipFile(path_b) as zb:
        files_a = {n for n in za.namelist() if not n.endswith("/")}
        files_b = {n for n in zb.namelist() if not n.endswith("/")}
        only_a = sorted(files_a - files_b)
        only_b = sorted(files_b - files_a)
        changed = []
        for name in sorted(files_a & files_b):
            hash_a = hashlib.sha256(za.read(name)).hexdigest()
            hash_b = hashlib.sha256(zb.read(name)).hexdigest()
            if hash_a != hash_b:
                changed.append(name)
        print("Comparison summary:")
        print(f"  Only in generated: {len(only_a)}")
        print(f"  Only in repaired:  {len(only_b)}")
        print(f"  Changed files:     {len(changed)}")
        if only_a:
            print("  Generated-only files:")
            for name in only_a:
                print(f"    - {name}")
        if only_b:
            print("  Repaired-only files:")
            for name in only_b:
                print(f"    - {name}")
        if changed:
            print("  Changed files:")
            for name in changed:
                print(f"    - {name}")
        if show_diff:
            for name in changed:
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                a_lines = read_text(za, name).splitlines()
                b_lines = read_text(zb, name).splitlines()
                diff = difflib.unified_diff(a_lines, b_lines, fromfile=f"generated:{name}", tofile=f"repaired:{name}", lineterm="")
                count = 0
                print(f"--- diff {name} ---")
                for line in diff:
                    print(line)
                    count += 1
                    if count >= diff_limit:
                        print("... diff truncated ...")
                        break


def check_relationship_targets(zf, rels_by_source, warnings):
    names = set(zf.namelist())
    for source, rels in rels_by_source.items():
        for rel in rels:
            if rel["target_mode"].lower() == "external":
                continue
            target = resolve_target(source, rel["target"])
            if target not in names:
                warnings.append(f"Missing target {target} for rel {rel['type']} in {source or '[package]'}")


def check_notes(zf, rels_by_source, overrides, warnings):
    names = set(zf.namelist())
    notes_slides = sorted(
        name for name in names if name.startswith("ppt/notesSlides/notesSlide") and name.endswith(".xml")
    )
    notes_masters = sorted(
        name for name in names if name.startswith("ppt/notesMasters/notesMaster") and name.endswith(".xml")
    )
    if notes_slides and not notes_masters:
        warnings.append("Notes slides exist but no notes master part found.")
    if notes_masters:
        pres_rels = rels_by_source.get("ppt/presentation.xml", [])
        if not any(rel["type"] == REL_TYPE_NOTES_MASTER for rel in pres_rels):
            warnings.append("Presentation relationships missing notesMaster relationship.")
        try:
            pres_xml = zf.read("ppt/presentation.xml")
            if b"notesMasterIdLst" not in pres_xml:
                warnings.append("presentation.xml missing notesMasterIdLst.")
        except KeyError:
            warnings.append("Missing ppt/presentation.xml")
        for master in notes_masters:
            part_name = f"/{master}"
            if part_name not in overrides:
                warnings.append(f"Content types missing override for {part_name}.")
        master_rels = rels_by_source.get(notes_masters[0], [])
        if not any(rel["type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" for rel in master_rels):
            warnings.append("Notes master missing theme relationship.")
    for slide in notes_slides:
        rels = rels_by_source.get(slide)
        if rels is None:
            warnings.append(f"Missing relationships for {slide}.")
            continue
        if not any(rel["type"] == REL_TYPE_NOTES_MASTER for rel in rels):
            warnings.append(f"{slide} missing notesMaster relationship.")
        if not any(rel["type"] == REL_TYPE_NOTES_SLIDE or rel["type"].endswith("/slide") for rel in rels):
            warnings.append(f"{slide} missing slide relationship.")
        part_name = f"/{slide}"
        if part_name not in overrides:
            warnings.append(f"Content types missing override for {part_name}.")


def check_comments(zf, rels_by_source, overrides, warnings):
    names = set(zf.namelist())
    comment_parts = sorted(
        name for name in names if name.startswith("ppt/comments/modernComment") and name.endswith(".xml")
    )
    if not comment_parts:
        return
    if "ppt/authors.xml" not in names:
        warnings.append("Comments exist but ppt/authors.xml is missing.")
    pres_rels = rels_by_source.get("ppt/presentation.xml", [])
    if not any(rel["type"] == REL_TYPE_AUTHORS for rel in pres_rels):
        warnings.append("Presentation relationships missing authors relationship.")
    comment_targets = set()
    for source, rels in rels_by_source.items():
        if not source.startswith("ppt/slides/"):
            continue
        for rel in rels:
            if rel["type"] == REL_TYPE_COMMENTS:
                comment_targets.add(resolve_target(source, rel["target"]))
    for comment in comment_parts:
        if comment not in comment_targets:
            warnings.append(f"Comment part {comment} is not referenced by any slide.")
        part_name = f"/{comment}"
        if part_name not in overrides:
            warnings.append(f"Content types missing override for {part_name}.")
    if "/ppt/authors.xml" not in overrides:
        warnings.append("Content types missing override for /ppt/authors.xml.")


def inspect(path, compare=None, show_diff=False, diff_limit=200):
    print(f"Inspecting PPTX: {path}")
    warnings = []
    with zipfile.ZipFile(path) as zf:
        overrides = {}
        try:
            overrides = parse_content_types(zf.read("[Content_Types].xml"))
        except KeyError:
            warnings.append("Missing [Content_Types].xml.")
        rels_by_source = load_relationships(zf)
        check_relationship_targets(zf, rels_by_source, warnings)
        check_notes(zf, rels_by_source, overrides, warnings)
        check_comments(zf, rels_by_source, overrides, warnings)
    if compare:
        diff_zips(path, compare, show_diff=show_diff, diff_limit=diff_limit)
    if warnings:
        print("Warnings:")
        for warning in warnings:
            print(f"  - {warning}")
        return 1
    print("No issues detected.")
    return 0


def main(argv):
    parser = argparse.ArgumentParser(description="Inspect PPTX structure and relationships.")
    parser.add_argument("pptx", help="Path to pptx file")
    parser.add_argument("--compare", help="Path to repaired pptx for diff")
    parser.add_argument("--diff", action="store_true", help="Show unified diffs for changed XML files")
    parser.add_argument("--diff-limit", type=int, default=200, help="Max diff lines per file")
    args = parser.parse_args(argv)
    return inspect(args.pptx, compare=args.compare, show_diff=args.diff, diff_limit=args.diff_limit)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
