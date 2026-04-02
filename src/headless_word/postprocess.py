"""Post-process saved .docx files to fix comment threading.

LO UNO supports setting ParaIdParent on annotations, but the OOXML export
reassigns internal paragraph IDs during serialization, breaking parent→child
linkage. This module patches the saved .docx ZIP to inject proper
commentsExtended.xml with w15:paraIdParent entries.
"""

from __future__ import annotations

import shutil
import xml.dom.minidom
import zipfile
from pathlib import Path


def fix_comment_threading(docx_path: str) -> bool:
    path = Path(docx_path)
    if not path.exists() or path.suffix != ".docx":
        return False

    with zipfile.ZipFile(path) as z:
        if "word/comments.xml" not in z.namelist():
            return False
        comments_raw = z.read("word/comments.xml")

    dom = xml.dom.minidom.parseString(comments_raw)
    comments = dom.getElementsByTagName("w:comment")
    if not comments:
        return False

    return False


def fix_comment_threading_with_state(
    docx_path: str,
    threading: list[dict],
) -> bool:
    """Fix comment threading in a saved docx using threading state from UNO.

    Args:
        docx_path: Path to the saved .docx file.
        threading: List of dicts with keys:
            - comment_id: int (w:comment w:id in comments.xml)
            - parent_comment_id: int | None (parent's w:id, or None if top-level)
            - resolved: bool

    Returns True if any threading was applied.
    """
    path = Path(docx_path)
    if not path.exists():
        return False

    # Filter to only comments that have threading or resolved status
    threaded = [
        t
        for t in threading
        if t.get("parent_comment_id") is not None or t.get("resolved", False)
    ]
    if not threaded:
        return False

    with zipfile.ZipFile(path) as z:
        all_files = {name: z.read(name) for name in z.namelist()}

    if "word/comments.xml" not in all_files:
        return False

    # Parse comments.xml
    dom = xml.dom.minidom.parseString(all_files["word/comments.xml"])
    comments = dom.getElementsByTagName("w:comment")

    # Build map: UNO order index -> comment id
    # The threading state from UNO is ordered by field enumeration,
    # and comment_id matches the w:id attribute in comments.xml
    comment_ids = set()
    for c in comments:
        comment_ids.add(int(c.getAttribute("w:id")))

    # Assign paraId to each comment's paragraph
    para_id_map = {}  # comment w:id -> paraId hex string
    for c in comments:
        cid = int(c.getAttribute("w:id"))
        pid = f"{cid + 1:08X}"
        para_id_map[cid] = pid
        paras = c.getElementsByTagName("w:p")
        if paras:
            paras[0].setAttribute("w14:paraId", pid)

    all_files["word/comments.xml"] = dom.toxml().encode("utf-8")

    # Build UNO ParaId -> comment w:id mapping
    # The threading state has para_id_parent which is the UNO ParaId of the parent.
    # We need to map UNO ParaId values to comment w:ids to find the right paraId.
    #
    # UNO ParaId values come from the PostItId counter — they are assigned by LO
    # internally and don't directly correspond to w:id. We need the caller to
    # provide the mapping.

    # Build commentsExtended.xml
    ext_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    ext_xml += '<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    ext_xml += 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15">\n'

    for t in threading:
        cid = t["comment_id"]
        if cid not in para_id_map:
            continue
        pid = para_id_map[cid]
        parent_cid = t.get("parent_comment_id")
        resolved = t.get("resolved", False)
        done = "1" if resolved else "0"

        if parent_cid is not None and parent_cid in para_id_map:
            parent_pid = para_id_map[parent_cid]
            ext_xml += f'  <w15:commentEx w15:paraId="{pid}" w15:paraIdParent="{parent_pid}" w15:done="{done}"/>\n'
        else:
            ext_xml += f'  <w15:commentEx w15:paraId="{pid}" w15:done="{done}"/>\n'

    ext_xml += "</w15:commentsEx>"
    all_files["word/commentsExtended.xml"] = ext_xml.encode("utf-8")

    # Ensure Content_Types.xml has the entry
    ct = all_files["[Content_Types].xml"].decode("utf-8")
    if "commentsExtended" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/commentsExtended.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>\n'
            "</Types>",
        )
        all_files["[Content_Types].xml"] = ct.encode("utf-8")

    # Ensure document.xml.rels has the relationship
    rels_key = "word/_rels/document.xml.rels"
    if rels_key in all_files:
        rels = all_files[rels_key].decode("utf-8")
        if "commentsExtended" not in rels:
            rels = rels.replace(
                "</Relationships>",
                '<Relationship Id="rIdCommentsEx" '
                'Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" '
                'Target="commentsExtended.xml"/>\n</Relationships>',
            )
            all_files[rels_key] = rels.encode("utf-8")

    # Write fixed docx
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in all_files.items():
            z.writestr(name, data)
    shutil.move(tmp, str(path))

    return True
