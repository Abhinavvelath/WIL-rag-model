"""
extract_and_chunk.py

Extracts headings, paragraphs, tagged callouts, image captions and tables
from the project's .docx knowledge base files and produces validated,
retrieval-ready chunks as JSON.

Design principles (see research notes in the accompanying conversation):
  1. Idempotent    - re-running on the same input file produces byte-identical
                     output (deterministic chunk IDs, no timestamps in content).
  2. Validated     - a hard validation gate runs after chunking and REFUSES to
                     write output if any integrity check fails. No silent
                     partial results.
  3. Traceable     - every chunk carries structured metadata (heading path,
                     applies_to, source, kind) rather than being bare text.
  4. Tested        - a small test suite (run with --test) locks in the
                     behaviours this script depends on: table atomicity,
                     image-caption atomicity, heading inheritance, and
                     re-run determinism.

Usage:
    python3 extract_and_chunk.py --test                 # run the test suite
    python3 extract_and_chunk.py input.docx out.json     # chunk one file
    python3 scripts/extract_and_chunk.py --all   # chunk the PDF and URL KBs
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import Optional

import docx

MIN_CHUNK_SIZE = 40          # chars; below this a chunk is almost certainly junk
MAX_CHUNK_SIZE = 1200        # chars; soft target ceiling for non-atomic chunks
TAG_RE = re.compile(r"^\[(WARNING|NOTE|TIP|FACT|DEFINITION)\]$")
IMAGE_RE = re.compile(r"^\[IMAGE:.*\]$")
TABLE_START_RE = re.compile(r"^\[TABLE START:")
TABLE_END_RE = re.compile(r"^\[TABLE END\]$")
APPLIES_RE = re.compile(r"^Applies to:\s*(.*)$")
SOURCE_RE = re.compile(r"^Sources?:\s*(.*)$")      # accepts "Source:" and "Sources:"


# --------------------------------------------------------------------------- schema
@dataclass
class Chunk:
    id: str                      # deterministic hash — same input -> same id
    doc_id: str                  # source filename, e.g. "pdf_kb_clean.docx"
    seq: int                     # position within the document (for ordering / debugging)
    kind: str                    # "text" | "table" | "image"
    heading_path: list[str]      # ["1. Driving a car in Victoria", "Toll roads"]
    applies_to: Optional[str]    # from the section's "Applies to:" line, if any
    source: Optional[str]        # from the section's "Source:" line, if any
    text: str                    # the raw chunk text
    embed_text: str              # text + heading breadcrumb, what actually gets embedded
    char_len: int = field(init=False)

    def __post_init__(self):
        self.char_len = len(self.text)


# --------------------------------------------------------------------------- extraction
def read_paragraphs(path: Path) -> list[str]:
    """Return one (style, line) tuple per LOGICAL line, not per Word paragraph object.

    Some of our source .docx files store a whole table or callout as a single
    Word paragraph with manual line breaks (python-docx then returns embedded
    '\\n' characters in that one paragraph's .text); others store one Word
    paragraph per line. Splitting every paragraph's text on '\\n' up front
    normalizes both cases to the same representation, so the pattern-matching
    below (TABLE_START/END, [WARNING] etc, "Applies to:") sees one line per
    unit regardless of which way the source file happened to be built.
    """
    d = docx.Document(str(path))
    out = []
    for p in d.paragraphs:
        raw = p.text.replace("\u00a0", " ").replace("\u200b", "")
        for line in raw.split("\n"):
            t = line.strip()
            if t:
                out.append((p.style.name, t))
    return out


def chunk_id(doc_id: str, kind: str, heading_path: list[str], text: str) -> str:
    """Deterministic ID: same doc + heading + text always yields the same id."""
    h = hashlib.sha1()
    h.update(doc_id.encode("utf-8"))
    h.update(b"\x00")
    h.update("/".join(heading_path).encode("utf-8"))
    h.update(b"\x00")
    h.update(text.encode("utf-8"))
    return h.hexdigest()[:16]


def build_chunks(path: Path) -> list[Chunk]:
    doc_id = path.name
    paras = read_paragraphs(path)
    chunks: list[Chunk] = []

    h1 = h2 = ""
    applies_to = None
    source = None
    h1_source = None      # the section-level source; a subsection's own Source line must not leak into the next subsection
    seq = 0
    buf: list[str] = []
    buf_len = 0

    def heading_path() -> list[str]:
        return [x for x in (h1, h2) if x]

    def flush():
        nonlocal buf, buf_len, seq
        if not buf:
            return
        text = "\n".join(buf).strip()
        if text:
            embed = ("\n".join(heading_path()) + "\n" + text) if heading_path() else text
            chunks.append(Chunk(
                id=chunk_id(doc_id, "text", heading_path(), text),
                doc_id=doc_id, seq=seq, kind="text",
                heading_path=heading_path(), applies_to=applies_to, source=source,
                text=text, embed_text=embed,
            ))
            seq += 1
        buf, buf_len = [], 0

    i = 0
    n = len(paras)
    while i < n:
        style, text = paras[i]

        if style in ("Heading 1", "Title"):
            flush(); h1, h2 = text, ""; applies_to = None; source = None; h1_source = None
            i += 1; continue
        if style == "Heading 2":
            flush(); h2 = text; source = h1_source
            i += 1; continue

        m = APPLIES_RE.match(text)
        if m:
            applies_to = m.group(1); i += 1; continue
        m = SOURCE_RE.match(text)
        if m:
            source = m.group(1)
            if not h2:
                h1_source = source
            i += 1; continue

        if TABLE_START_RE.match(text):
            flush()
            j = i
            block = [text]
            j += 1
            while j < n and not TABLE_END_RE.match(paras[j][1]):
                block.append(paras[j][1]); j += 1
            if j < n:
                block.append(paras[j][1]); j += 1  # include the END line
            full = "\n".join(block)
            embed = ("\n".join(heading_path()) + "\n" + full) if heading_path() else full
            chunks.append(Chunk(
                id=chunk_id(doc_id, "table", heading_path(), full),
                doc_id=doc_id, seq=seq, kind="table",
                heading_path=heading_path(), applies_to=applies_to, source=source,
                text=full, embed_text=embed,
            ))
            seq += 1
            i = j
            continue

        if IMAGE_RE.match(text) or TAG_RE.match(text):
            flush()
            unit = [text]
            j = i + 1
            if j < n and (paras[j][1].startswith("Meaning for the driver:") or TAG_RE.match(text)):
                unit.append(paras[j][1]); j += 1
            full = "\n".join(unit)
            kind = "image" if IMAGE_RE.match(text) else "text"
            embed = ("\n".join(heading_path()) + "\n" + full) if heading_path() else full
            chunks.append(Chunk(
                id=chunk_id(doc_id, kind, heading_path(), full),
                doc_id=doc_id, seq=seq, kind=kind,
                heading_path=heading_path(), applies_to=applies_to, source=source,
                text=full, embed_text=embed,
            ))
            seq += 1
            i = j
            continue

        # ordinary paragraph: accumulate up to the soft size limit
        buf.append(text); buf_len += len(text)
        if buf_len > MAX_CHUNK_SIZE:
            flush()
        i += 1

    flush()
    return chunks


# --------------------------------------------------------------------------- validation gate
class ValidationError(Exception):
    pass


def validate(chunks: list[Chunk], doc_id: str, min_total_chars: int = 1000) -> list[str]:
    """Returns a list of human-readable problems. Empty list = pass.

    min_total_chars guards against silent data loss on a full document run;
    it is lowered for small synthetic fixtures in the test suite, which are
    deliberately tiny and would otherwise trip this check for the wrong reason.
    """
    problems = []

    if not chunks:
        problems.append(f"{doc_id}: produced zero chunks")
        return problems

    # 1. no chunk is empty / whitespace-only / absurdly small (except atomic table/image which may be short)
    tiny = [c for c in chunks if c.char_len < MIN_CHUNK_SIZE and c.kind == "text"]
    if tiny:
        problems.append(f"{doc_id}: {len(tiny)} text chunk(s) under {MIN_CHUNK_SIZE} chars "
                         f"(example id={tiny[0].id}: {tiny[0].text[:60]!r})")

    # 1b. no *text* chunk should ever be wildly larger than the soft target — a huge
    #     text chunk almost always means heading/flush detection silently failed
    #     upstream (e.g. this caught a real bug: a table packed into one Word
    #     paragraph swallowed the rest of the document because no TABLE END line
    #     was ever found as its own unit).
    huge = [c for c in chunks if c.kind == "text" and c.char_len > MAX_CHUNK_SIZE * 5]
    if huge:
        problems.append(f"{doc_id}: {len(huge)} text chunk(s) over {MAX_CHUNK_SIZE * 5} chars "
                         f"(example id={huge[0].id}, {huge[0].char_len} chars) — "
                         f"likely means a heading or table boundary was missed upstream")

    # 2. every table chunk must contain both START and END markers (i.e. not split)
    for c in chunks:
        if c.kind == "table":
            if "[TABLE START" not in c.text or "[TABLE END]" not in c.text:
                problems.append(f"{doc_id}: table chunk {c.id} missing START or END marker")

    # 3. every image chunk must carry its meaning line (not orphaned)
    for c in chunks:
        if c.kind == "image" and "Meaning for the driver:" not in c.text:
            problems.append(f"{doc_id}: image chunk {c.id} has no 'Meaning for the driver' line "
                             f"({c.text[:60]!r})")

    # 4. no duplicate chunk IDs (would indicate a hashing collision or duplicate content)
    ids = [c.id for c in chunks]
    dupes = {x for x in ids if ids.count(x) > 1}
    if dupes:
        problems.append(f"{doc_id}: {len(dupes)} duplicate chunk id(s): {list(dupes)[:3]}")

    # 5. every chunk has a non-empty heading path (except possibly the very first chunk, which we allow to warn only)
    no_heading = [c for c in chunks if not c.heading_path]
    if len(no_heading) > 1:  # allow at most 1 (e.g. a title-page preamble)
        problems.append(f"{doc_id}: {len(no_heading)} chunks with no heading path "
                         f"(example id={no_heading[0].id})")

    # 6. total reconstructed text should roughly match source length (sanity check against silent data loss)
    total_chars = sum(c.char_len for c in chunks)
    if total_chars < min_total_chars:
        problems.append(f"{doc_id}: suspiciously small total output ({total_chars} chars, expected >= {min_total_chars})")

    return problems


# --------------------------------------------------------------------------- driver
def process(path: Path, out_path: Path) -> None:
    chunks = build_chunks(path)
    problems = validate(chunks, path.name)
    if problems:
        msg = f"\nValidation FAILED for {path.name} — output NOT written.\n" + "\n".join(f"  - {p}" for p in problems)
        raise ValidationError(msg)

    kinds = {}
    for c in chunks:
        kinds[c.kind] = kinds.get(c.kind, 0) + 1
    lens = [c.char_len for c in chunks]

    out_path.write_text(
        json.dumps([asdict(c) for c in chunks], indent=1, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"OK  {path.name} -> {out_path.name}")
    print(f"    chunks: {len(chunks)}  ({', '.join(f'{k}={v}' for k, v in sorted(kinds.items()))})")
    print(f"    chunk length: min={min(lens)} median={sorted(lens)[len(lens)//2]} max={max(lens)}")
    print(f"    validation: all {7} checks passed")


# --------------------------------------------------------------------------- tests
def _mk_tmp_docx(path: Path, paragraphs: list[tuple[str, str]]):
    d = docx.Document()
    for style, text in paragraphs:
        if style == "Heading 1":
            d.add_heading(text, level=1)
        elif style == "Heading 2":
            d.add_heading(text, level=2)
        else:
            d.add_paragraph(text)
    d.save(str(path))


def run_tests() -> bool:
    import tempfile
    ok = True

    def check(name, cond):
        nonlocal ok
        status = "PASS" if cond else "FAIL"
        if not cond:
            ok = False
        print(f"  [{status}] {name}")

    with tempfile.TemporaryDirectory() as td:
        tdp = Path(td)

        # --- Test 1: table atomicity -------------------------------------------------
        p1 = tdp / "t1.docx"
        _mk_tmp_docx(p1, [
            ("Heading 1", "Fines"),
            ("Normal", "Applies to: all drivers"),
            ("Normal", "Source: test"),
            ("Normal", "[TABLE START: Speeding]"),
            ("Normal", "| Speed | Fine |"),
            ("Normal", "|---|---|"),
            ("Normal", "| 10 km/h | $254 |"),
            ("Normal", "[TABLE END]"),
            ("Normal", "Some trailing paragraph long enough to clear the minimum chunk size floor on its own."),
        ])
        chunks = build_chunks(p1)
        table_chunks = [c for c in chunks if c.kind == "table"]
        check("table produces exactly one atomic chunk", len(table_chunks) == 1)
        check("table chunk contains START and END", table_chunks and "[TABLE START" in table_chunks[0].text and "[TABLE END]" in table_chunks[0].text)
        check("table chunk carries applies_to metadata", table_chunks and table_chunks[0].applies_to == "all drivers")
        problems = validate(chunks, "t1.docx", min_total_chars=10)  # tiny fixture; real docs use the default 1000
        check("validation passes on well-formed table doc", problems == [])

        # --- Test 2: image caption atomicity -----------------------------------------
        p2 = tdp / "t2.docx"
        _mk_tmp_docx(p2, [
            ("Heading 1", "Signs"),
            ("Normal", "[IMAGE: Hook turn sign]"),
            ("Normal", "Meaning for the driver: turn from the far-left lane."),
            ("Normal", "Some following paragraph long enough to be a real chunk on its own merit here."),
        ])
        chunks = build_chunks(p2)
        img_chunks = [c for c in chunks if c.kind == "image"]
        check("image produces exactly one atomic chunk", len(img_chunks) == 1)
        check("image chunk carries its meaning line", img_chunks and "Meaning for the driver" in img_chunks[0].text)
        problems = validate(chunks, "t2.docx", min_total_chars=10)  # tiny fixture
        check("validation passes on well-formed image doc", problems == [])

        # --- Test 3: heading inheritance -----------------------------------------------
        p3 = tdp / "t3.docx"
        _mk_tmp_docx(p3, [
            ("Heading 1", "Section A"),
            ("Heading 2", "Subsection A1"),
            ("Normal", "Some paragraph text that belongs under subsection A1, long enough to pass the size floor."),
        ])
        chunks = build_chunks(p3)
        check("heading path captured correctly", chunks and chunks[0].heading_path == ["Section A", "Subsection A1"])

        # --- Test 3b: plural "Sources:" line is read as metadata; a subsection's own source does not leak -------
        p3b = tdp / "t3b.docx"
        _mk_tmp_docx(p3b, [
            ("Heading 1", "Section"),
            ("Normal", "Applies to: all drivers"),
            ("Normal", "Sources: the section source"),
            ("Heading 2", "First subsection"),
            ("Normal", "Source: only for the first subsection"),
            ("Normal", "Text in the first subsection, long enough to be a real chunk on its own merit here."),
            ("Heading 2", "Second subsection"),
            ("Normal", "Text in the second subsection, long enough to be a real chunk on its own merit here too."),
        ])
        chunks = build_chunks(p3b)
        by = {c.heading_path[-1]: c.source for c in chunks}
        check("plural 'Sources:' line is read as source metadata", by.get("First subsection") == "only for the first subsection")
        check("a subsection's own source does not leak into the next subsection", by.get("Second subsection") == "the section source")

        # --- Test 4: validation gate catches a broken table (simulated corruption) -----
        broken = [Chunk(id="x", doc_id="broken.docx", seq=0, kind="table",
                         heading_path=["A"], applies_to=None, source=None,
                         text="[TABLE START: X]\n| a | b |", embed_text="x")]
        problems = validate(broken, "broken.docx")
        check("validation gate catches a table missing its END marker", any("missing START or END" in p for p in problems))

        # --- Test 5: validation gate catches an orphaned image caption -----------------
        orphan = [Chunk(id="y", doc_id="broken.docx", seq=0, kind="image",
                         heading_path=["A"], applies_to=None, source=None,
                         text="[IMAGE: test sign]", embed_text="x")]
        problems = validate(orphan, "broken.docx")
        check("validation gate catches an orphaned image caption", any("no 'Meaning for the driver'" in p for p in problems))

        # --- Test 5b: regression test for the "table packed into one Word paragraph
        #     with embedded line breaks" bug found when this script was first run on
        #     the real pdf_kb_clean.docx file. A single python-docx Paragraph whose
        #     .text already contains '\n' characters (as produced by run.add_break())
        #     must still be split into individual TABLE START/END lines.
        p5b = tdp / "t5b.docx"
        dd = docx.Document()
        dd.add_heading("Regression: multi-line paragraph", level=1)
        one_para = dd.add_paragraph()
        lines = ["[TABLE START: X]", "| a | b |", "|---|---|", "| 1 | 2 |", "[TABLE END]"]
        for idx, ln in enumerate(lines):
            r = one_para.add_run(ln)
            if idx < len(lines) - 1:
                r.add_break()
        dd.save(str(p5b))
        chunks = build_chunks(p5b)
        table_chunks = [c for c in chunks if c.kind == "table"]
        check("table packed into a single Word paragraph (embedded line breaks) is still parsed as one atomic table",
              len(table_chunks) == 1 and "[TABLE END]" in table_chunks[0].text)
        problems = validate(chunks, "t5b.docx", min_total_chars=10)
        check("validation passes on the embedded-linebreak table doc", problems == [])

        # --- Test 6: idempotency / determinism -----------------------------------------
        p6 = tdp / "t6.docx"
        _mk_tmp_docx(p6, [
            ("Heading 1", "Repeatable section"),
            ("Normal", "This paragraph should hash identically every single time we build chunks from it."),
        ])
        run1 = build_chunks(p6)
        run2 = build_chunks(p6)
        ids1 = [c.id for c in run1]
        ids2 = [c.id for c in run2]
        check("re-running on identical input yields identical chunk ids", ids1 == ids2 and len(ids1) > 0)

        # --- Test 7: process() refuses to write output when validation fails -----------
        p7 = tdp / "t7.docx"
        _mk_tmp_docx(p7, [("Heading 1", "Only a heading, no body text at all")])
        out7 = tdp / "t7.json"
        raised = False
        try:
            process(p7, out7)
        except ValidationError:
            raised = True
        check("process() raises and writes nothing when a doc has no real content",
              raised and not out7.exists())

    return ok


# --------------------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", nargs="?", help="input .docx path")
    ap.add_argument("output", nargs="?", help="output .json path")
    ap.add_argument("--all", action="store_true", help="process the PDF and URL knowledge bases")
    ap.add_argument("--test", action="store_true", help="run the test suite and exit")
    args = ap.parse_args()

    if args.test:
        print("Running test suite for extract_and_chunk.py\n")
        ok = run_tests()
        print("\nALL TESTS PASSED" if ok else "\nSOME TESTS FAILED")
        sys.exit(0 if ok else 1)

    if args.all:
        base = Path(__file__).resolve().parent
        targets = [
            (base.parent / "knowledge-base" / "PDF_KB_self_contained.docx", base.parent / "data" / "pdf_kb_chunks.json"),
            (base.parent / "knowledge-base" / "URL_KB_self_contained.docx", base.parent / "data" / "url_kb_chunks.json"),
        ]
        failed = False
        for inp, outp in targets:
            try:
                process(inp, outp)
            except ValidationError as e:
                print(e)
                failed = True
        sys.exit(1 if failed else 0)

    if args.input and args.output:
        try:
            process(Path(args.input), Path(args.output))
        except ValidationError as e:
            print(e)
            sys.exit(1)
        return

    ap.print_help()


if __name__ == "__main__":
    main()
