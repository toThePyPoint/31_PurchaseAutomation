#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations
import logging 
import time
from typing import Dict, List, Tuple, Sequence, Optional, Any
from pyrfc import Connection, CommunicationError, LogonError, ABAPApplicationError, ABAPRuntimeError
from .log_utils import setup_logger

log = setup_logger("SAP_RTAB", "sap_rtab.log")

MAX_OPT = 72

def chunk_list(lst: List[Any], size: int) -> List[List[Any]]:
    return [lst[i : i + size] for i in range(0, len(lst), size)]

def split_where(where: str, max_len: int = MAX_OPT) -> List[str]:
    parts: List[str] = []
    s = where.strip()
    while len(s) > max_len:
        cut_and = s.rfind(" AND ", 0, max_len)
        cut_or = s.rfind(" OR ", 0, max_len)
        cut = max(cut_and, cut_or)
        if cut <= 0:
            parts.append(s[:max_len])
            s = s[max_len:]
        else:
            parts.append(s[:cut].strip())
            s = s[cut + 1 :].strip()
    if s:
        parts.append(s)
    return [p for p in parts if p]

def options_from_where(where: str) -> List[Dict[str, str]]:
    where = (where or "").strip()
    if not where:
        return []
    return [{"TEXT": ln} for ln in split_where(where)]

def _options_from_lines(where_lines: Sequence[str]) -> List[Dict[str, str]]:
    parts: List[str] = []
    for i, raw in enumerate(where_lines):
        t = raw.strip()
        if not t: continue
        if i > 0 and not t.upper().startswith(("AND ", "OR ")):
            t = "AND " + t
        parts.append(t)
    where = " ".join(parts)
    return options_from_where(where)

def rfc_read_table(
    conn: Connection,
    table: str,
    fields: Sequence[str],
    where: str = "",
    rowcount: int = 0,
    rowskips: int = 0,
    delimiter: str = "§",
) -> List[Dict[str, str]]:

    try:
        res = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER=delimiter,
            FIELDS=[{"FIELDNAME": f} for f in fields],
            OPTIONS=options_from_where(where),
            ROWCOUNT=rowcount,
            ROWSKIPS=rowskips,
        )
    except (CommunicationError, LogonError, ABAPApplicationError, ABAPRuntimeError) as e:
        log.error("RFC_READ_TABLE error on %s: %s", table, e)
        raise

    cols = [f["FIELDNAME"] for f in res.get("FIELDS", [])]
    out: List[Dict[str, str]] = []
    for row in res.get("DATA", []):
        wa = row.get("WA", "")
        parts = wa.split(delimiter)
        if len(parts) < len(cols):
            parts += [""] * (len(cols) - len(parts))
        out.append({c: p.strip() for c, p in zip(cols, parts)})
    return out

def rtab(
    conn: Connection,
    table: str,
    fields: Sequence[str],
    where_lines: Sequence[str],
    per_page: int = 20000,
    max_rows: int = 500_000,
    delimiter: str = "§",
    logger: Optional[logging.Logger] = None,
    max_retries: int = 3,
    retry_delay: int = 5
) -> List[str]:

    rows: List[str] = []
    fields_def = [{"FIELDNAME": f} for f in fields]
    opts = _options_from_lines(where_lines)
    skips = 0
    lg = logger or log

    lg.debug("RFC_READ_TABLE %s: lines=%d", table, len(where_lines))

    while True:
        chunk_data = []
        
        for attempt in range(1, max_retries + 1):
            try:
                resp = conn.call(
                    "RFC_READ_TABLE",
                    QUERY_TABLE=table,
                    DELIMITER=delimiter,
                    FIELDS=fields_def,
                    OPTIONS=opts,
                    ROWCOUNT=per_page,
                    ROWSKIPS=skips,
                )
                chunk_data = resp.get("DATA", [])
                break # Success
            except (CommunicationError, LogonError) as e:
                lg.warning(f"Błąd sieci przy {table} (próba {attempt}/{max_retries}): {e}")
                if attempt == max_retries:
                    lg.error(f"Wyczerpano limit prób dla {table}.")
                    raise e
                time.sleep(retry_delay)
            except Exception as e:
                lg.error(f"Krytyczny błąd RFC przy {table}: {e}")
                raise e

        if not chunk_data:
            break

        rows.extend(d.get("WA", "") for d in chunk_data)

        if len(rows) >= max_rows:
            lg.info(f"Osiągnięto limit wierszy ({max_rows}) dla {table}")
            rows = rows[:max_rows]
            break

        if len(chunk_data) < per_page:
            break

        skips += per_page

    return rows