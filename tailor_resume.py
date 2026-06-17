#!/usr/bin/env python3
"""
ATS Resume Tailor (Gemini Edition)
------------------------------------
Takes a .docx resume + job description, uses Google Gemini AI (free tier)
to tailor content for ATS keyword matching, then writes back into the
original Word document preserving ALL formatting.

Setup:
    pip install google-genai lxml python-docx
    export GEMINI_API_KEY=your-key-here   # free at aistudio.google.com

Usage:
    python3 tailor_resume.py --resume resume.docx --jd job_description.txt --output tailored_resume.docx
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
from pathlib import Path
from lxml import etree
from google import genai
from google.genai import types

# ── Namespaces used in OOXML ──────────────────────────────────────────────────
NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}

# Path to docx skill scripts (only available in Claude environment)
SKILLS_SCRIPTS = Path("/mnt/skills/public/docx/scripts/office")


# ── Step 1 – Unpack the docx ─────────────────────────────────────────────────
def unpack_docx(docx_path: Path, unpack_dir: Path) -> None:
    """Unpack docx into XML — tries skill script first, falls back to unzip."""
    if unpack_dir.exists():
        shutil.rmtree(unpack_dir)

    if SKILLS_SCRIPTS.exists():
        result = subprocess.run(
            ["python3", str(SKILLS_SCRIPTS / "unpack.py"), str(docx_path), str(unpack_dir)],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print(f"[✓] Unpacked to {unpack_dir}")
            return
        print("[!] Skill unpack failed, falling back to unzip...")

    # Fallback: standard unzip (available on macOS and Linux)
    unpack_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(["unzip", "-q", str(docx_path), "-d", str(unpack_dir)], check=True)
    print(f"[✓] Unpacked to {unpack_dir}")


# ── Step 2 – Extract readable text from document.xml ─────────────────────────
def extract_text_from_xml(doc_xml: Path) -> str:
    """Extract all paragraph text from OOXML, preserving line structure."""
    tree = etree.parse(str(doc_xml))
    root = tree.getroot()

    paragraphs = []
    for para in root.iter(f"{{{NS['w']}}}p"):
        runs = para.findall(f".//{{{NS['w']}}}t")
        text = "".join(r.text or "" for r in runs).strip()
        if text:
            paragraphs.append(text)

    return "\n".join(paragraphs)


# ── Step 3 – Tailor with Gemini ──────────────────────────────────────────────
def tailor_with_gemini(resume_text: str, jd_text: str) -> dict:
    """
    Send resume + JD to Gemini. Returns:
      {
        "replacements": [{"original": "...", "replacement": "..."}, ...],
        "changes_summary": ["change 1", ...]
      }
    Uses the new google-genai SDK (replaces deprecated google-generativeai).
    """
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY not set.\n"
            "  1. Get a free key at https://aistudio.google.com/app/apikey\n"
            "  2. Run:  export GEMINI_API_KEY=your-key-here"
        )

    client = genai.Client(api_key=api_key)

    prompt = textwrap.dedent(f"""
        You are an expert ATS resume optimizer. Tailor the resume below to match
        the job description for maximum ATS keyword score — without inventing experience.

        RULES:
        1. Only suggest replacements for text that EXISTS verbatim in the resume.
        2. Replacements must be similar in length (±20%) to preserve single-page layout.
        3. Incorporate keywords, technologies, and role-specific language from the JD naturally.
        4. Do NOT change: names, contact info, dates, company names, job titles, section headings.
        5. Use strong action verbs and measurable impact language that mirrors the JD tone.
        6. Keep the resume to ONE page — be concise, trim verbose phrases.
        7. Return ONLY valid JSON — no markdown, no explanation, no backticks.

        JSON format:
        {{
          "replacements": [
            {{"original": "exact text from resume", "replacement": "tailored replacement"}},
            ...
          ],
          "changes_summary": [
            "Brief description of each change made",
            ...
          ]
        }}

        Skip any sentence already well-optimized for the JD.

        RESUME:
        {resume_text}

        JOB DESCRIPTION:
        {jd_text}
    """).strip()

    print("[⏳] Sending to Gemini AI for tailoring...")
    response = client.models.generate_content(
        model="gemini-3.1-flash-lite",
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.3,
            response_mime_type="application/json",
        ),
    )

    raw = response.text.strip()
    # Safety net: strip any accidental markdown fences
    raw = re.sub(r"^```(?:json)?", "", raw, flags=re.MULTILINE).strip()
    raw = re.sub(r"```$",          "", raw, flags=re.MULTILINE).strip()

    try:
        result = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"[!] Gemini returned invalid JSON: {e}", file=sys.stderr)
        print(f"Raw response (first 500 chars):\n{raw[:500]}", file=sys.stderr)
        raise

    print(f"[✓] Gemini returned {len(result.get('replacements', []))} replacements")
    return result


# ── Step 4 – Apply replacements to document.xml ──────────────────────────────
def apply_replacements_to_xml(doc_xml: Path, replacements: list[dict]) -> int:
    """
    Apply text replacements directly into the XML string.
    All run-level formatting (<w:rPr>) stays completely untouched.
    Returns number of replacements successfully applied.
    """
    xml_content = doc_xml.read_text(encoding="utf-8")
    applied = 0
    skipped = []

    for item in replacements:
        original    = item.get("original",    "").strip()
        replacement = item.get("replacement", "").strip()

        if not original or not replacement or original == replacement:
            continue

        if original in xml_content:
            xml_content = xml_content.replace(original, replacement, 1)
            applied += 1
            print(f"  [✓] {original[:60]}...")
        else:
            skipped.append(original[:60])

    doc_xml.write_text(xml_content, encoding="utf-8")

    if skipped:
        print(f"\n[!] {len(skipped)} replacements not found "
              f"(text may be split across XML runs):")
        for s in skipped[:5]:
            print(f"    • {s}...")
        if len(skipped) > 5:
            print(f"    ... and {len(skipped) - 5} more")

    return applied


# ── Step 5 – Repack into docx ────────────────────────────────────────────────
def pack_docx(unpack_dir: Path, output_path: Path, original_docx: Path) -> None:
    """Repack the edited XML back into a .docx."""
    if SKILLS_SCRIPTS.exists():
        result = subprocess.run(
            ["python3", str(SKILLS_SCRIPTS / "pack.py"),
             str(unpack_dir), str(output_path),
             "--original", str(original_docx)],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print(f"[✓] Packed to {output_path}")
            return
        print("[!] Skill pack failed, falling back to zip repack...")

    # Fallback: plain zip repack
    output_path.unlink(missing_ok=True)
    subprocess.run(
        f'cd "{unpack_dir}" && zip -qr "{output_path.resolve()}" .',
        shell=True, check=True
    )
    print(f"[✓] Packed to {output_path}")


# ── Helpers ───────────────────────────────────────────────────────────────────
def read_file_safe(path: Path) -> str:
    """Read a text file, auto-detecting encoding to handle UTF-8, UTF-16, etc."""
    # Try encodings in order of likelihood
    for encoding in ("utf-8", "utf-16", "latin-1", "cp1252"):
        try:
            text = path.read_text(encoding=encoding)
            print(f"[✓] Loaded {path.name} ({encoding})")
            return text
        except (UnicodeDecodeError, UnicodeError):
            continue
    # Last resort: ignore undecodable bytes
    text = path.read_text(encoding="utf-8", errors="ignore")
    print(f"[!] Loaded {path.name} with some characters ignored (unknown encoding)")
    return text


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Tailor a .docx resume to a job description using Gemini AI (free)"
    )
    parser.add_argument("--resume", required=True, help="Path to your resume .docx")
    parser.add_argument("--jd",     required=True, help="Job description: .txt file path OR raw text")
    parser.add_argument("--output", default="tailored_resume.docx", help="Output .docx path")
    args = parser.parse_args()

    resume_path = Path(args.resume)
    output_path = Path(args.output)

    if not resume_path.exists():
        print(f"[✗] Resume not found: {resume_path}", file=sys.stderr)
        sys.exit(1)

    # Read JD — accept file path or raw inline text
    jd_path = Path(args.jd)
    if jd_path.exists():
        jd_text = read_file_safe(jd_path)
    else:
        jd_text = args.jd
        print(f"[✓] Using inline JD ({len(jd_text)} chars)")

    if len(jd_text.strip()) < 100:
        print("[✗] JD too short — paste at least 100 characters.", file=sys.stderr)
        sys.exit(1)

    with tempfile.TemporaryDirectory() as tmpdir:
        unpack_dir = Path(tmpdir) / "unpacked"
        doc_xml    = unpack_dir / "word" / "document.xml"

        print(f"\n[1/5] Unpacking {resume_path.name}...")
        unpack_docx(resume_path, unpack_dir)

        print("[2/5] Extracting resume text...")
        resume_text = extract_text_from_xml(doc_xml)
        word_count  = len(resume_text.split())
        print(f"[✓] Extracted {word_count} words")
        if word_count < 50:
            print("[!] Warning: very few words extracted — resume may use text boxes or images.")

        print("[3/5] Tailoring with Gemini AI...")
        result          = tailor_with_gemini(resume_text, jd_text)
        replacements    = result.get("replacements", [])
        changes_summary = result.get("changes_summary", [])

        print(f"[4/5] Applying {len(replacements)} replacements...")
        applied = apply_replacements_to_xml(doc_xml, replacements)

        print("[5/5] Repacking .docx...")
        pack_docx(unpack_dir, output_path, resume_path)

    pad = 32 - len(str(applied))
    print(f"""
╔══════════════════════════════════════════════════════╗
║              ATS Tailoring Complete!                 ║
╠══════════════════════════════════════════════════════╣
║  Input:    {resume_path.name:<41}║
║  Output:   {output_path.name:<41}║
║  Changes:  {applied} replacements applied{' ' * pad}║
╚══════════════════════════════════════════════════════╝

Changes made:""")
    for i, change in enumerate(changes_summary, 1):
        print(f"  {i}. {textwrap.fill(change, 65, subsequent_indent='     ')}")

    print(f"\n[✓] Ready: {output_path.resolve()}")


if __name__ == "__main__":
    main()
