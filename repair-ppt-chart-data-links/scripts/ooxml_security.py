# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

"""Security limits and XML parsing helpers for OOXML ZIP packages."""

from __future__ import annotations

import re
import zipfile
from pathlib import PurePosixPath

from lxml import etree


MAX_ARCHIVE_MEMBERS = 10_000
MAX_MEMBER_UNCOMPRESSED_BYTES = 64 * 1024 * 1024
MAX_TOTAL_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_COMPRESSION_RATIO = 1_000


class OoxmlSecurityError(RuntimeError):
    """Raised when an OOXML package exceeds safe local-processing limits."""


def validate_ooxml_archive(archive: zipfile.ZipFile) -> None:
    """Reject encrypted, path-unsafe, or decompression-risk ZIP members."""

    members = archive.infolist()
    if len(members) > MAX_ARCHIVE_MEMBERS:
        raise OoxmlSecurityError(
            f"OOXML archive has too many members: {len(members)}"
        )

    total_expanded = 0
    for member in members:
        normalized = member.filename.replace("\\", "/")
        path = PurePosixPath(normalized)
        if (
            not normalized
            or normalized.startswith("/")
            or re.match(r"^[A-Za-z]:", normalized)
            or ".." in path.parts
        ):
            raise OoxmlSecurityError(f"Unsafe OOXML member path: {member.filename}")
        if member.flag_bits & 0x1:
            raise OoxmlSecurityError(
                f"Encrypted OOXML member is not supported: {member.filename}"
            )
        if member.file_size > MAX_MEMBER_UNCOMPRESSED_BYTES:
            raise OoxmlSecurityError(
                f"OOXML member expanded size exceeds limit: {member.filename}"
            )
        total_expanded += member.file_size
        if total_expanded > MAX_TOTAL_UNCOMPRESSED_BYTES:
            raise OoxmlSecurityError("OOXML archive total expanded size exceeds limit")
        if member.file_size:
            if member.compress_size <= 0:
                raise OoxmlSecurityError(
                    f"OOXML member has invalid compressed size: {member.filename}"
                )
            ratio = member.file_size / member.compress_size
            if ratio > MAX_COMPRESSION_RATIO:
                raise OoxmlSecurityError(
                    f"OOXML member compression ratio exceeds limit: {member.filename}"
                )


def safe_xml_fromstring(payload: bytes | str) -> etree._Element:
    """Parse OOXML without DTDs, entities, network access, or huge trees."""

    raw = payload.encode("utf-8") if isinstance(payload, str) else payload
    probe = raw[:4096].upper()
    if b"<!DOCTYPE" in probe or b"<!ENTITY" in probe:
        raise OoxmlSecurityError("OOXML XML contains a prohibited DTD or entity")
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        huge_tree=False,
        recover=False,
    )
    try:
        return etree.fromstring(raw, parser=parser)
    except etree.XMLSyntaxError as exc:
        raise OoxmlSecurityError(f"Invalid OOXML XML: {exc}") from exc
