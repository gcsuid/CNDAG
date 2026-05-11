import base64
import json
import os

def base64_txt_to_json(input_file, output_file=None):

    with open(input_file, "r") as file:
        base64_str = file.read().strip()

    decoded_str = base64.b64decode(base64_str).decode("utf-8")


    json_data = json.loads(decoded_str)


    if output_file:
        with open(output_file, "w") as outfile:
            json.dump(json_data, outfile, indent=4)

    return json_data



def extract_cndag_jobs(json_data):
    jobs = []

    for section in json_data:
        for content in section.get("content", []):
            for field in content.get("fields", []):
                for condition in field.get("conditions", []):
                    func = condition.get("function", {})
                    for arg in func.get("argList", []):
                        ds = arg.get("dataSource", {})
                        if ds.get("type") == "CNDAG PIPELINE":
                            jobs.append({
                                "section_name": section["sectionName"],
                                "section_search": ds.get(
                                    "sectionSearch",
                                    section["sectionName"]
                                ),
                                "ds_id": ds.get("dsId")
                            })

    return jobs


from docx import Document

def clone_paragraph(src_paragraph, dest_doc):

    dst_para = dest_doc.add_paragraph()
    dst_para.style = src_paragraph.style
    dst_para.paragraph_format.alignment = src_paragraph.paragraph_format.alignment

    for run in src_paragraph.runs:
        new_run = dst_para.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        new_run.font.name = run.font.name
        new_run.font.size = run.font.size

import shutil

def copy_reference_doc(reference_doc_path, output_doc_path):
    shutil.copy(reference_doc_path, output_doc_path)


from docx import Document
from copy import deepcopy
import os

from docx import Document

from docx import Document

from docx import Document

def keep_only_section(doc_path: str, section_search: str):
    doc = Document(doc_path)
    body = doc.element.body
    blocks = list(body)

    # Map paragraph XML -> paragraph index
    para_map = {
        p._element: p for p in doc.paragraphs
    }

    start_idx = None
    start_level = None

    # Find start index and heading level
    for i, block in enumerate(blocks):
        para = para_map.get(block)
        if para:
            text = para.text.strip()
            if section_search.lower() in text.lower():
                style = para.style.name
                if style.startswith("Heading"):
                    start_level = int(style.split()[-1])
                    start_idx = i
                    break

    # If section not found, leave document unchanged
    if start_idx is None:
        doc.save(doc_path)
        return

    # Find end index
    end_idx = len(blocks)

    for i in range(start_idx + 1, len(blocks)):
        para = para_map.get(blocks[i])
        if para and para.style.name.startswith("Heading"):
            level = int(para.style.name.split()[-1])
            if level <= start_level:
                end_idx = i
                break

    # Remove everything
    for block in blocks:
        body.remove(block)

    # Restore only the section range
    for block in blocks[start_idx:end_idx]:
        body.append(block)

    doc.save(doc_path)

def keep_only_multiple_sections(doc_path: str, section_names: list[str]):
    doc = Document(doc_path)
    body = doc.element.body
    blocks = list(body)

    para_map = {p._element: p for p in doc.paragraphs}
    section_names = [s.lower() for s in section_names]

    keep_ranges = []
    i = 0

    while i < len(blocks):
        para = para_map.get(blocks[i])
        if para:
            text = para.text.strip().lower()
            if any(s in text for s in section_names) and para.style.name.startswith("Heading"):
                start = i
                base_level = int(para.style.name.split()[-1])

                end = len(blocks)
                for j in range(i + 1, len(blocks)):
                    p2 = para_map.get(blocks[j])
                    if p2 and p2.style.name.startswith("Heading"):
                        lvl = int(p2.style.name.split()[-1])
                        if lvl <= base_level:
                            end = j
                            break

                keep_ranges.append((start, end))
                i = end
                continue
        i += 1

    kept_blocks = []
    for start, end in keep_ranges:
        kept_blocks.extend(blocks[start:end])

    for block in blocks:
        body.remove(block)

    for block in kept_blocks:
        body.append(block)

    doc.save(doc_path)

def generate_final_cndag_doc(
    json_data,
    reference_doc_path,
    output_doc_path
):
    jobs = extract_cndag_jobs(json_data)
    section_names = [job["section_search"] for job in jobs]

    copy_reference_doc(reference_doc_path, output_doc_path)
    keep_only_multiple_sections(output_doc_path, section_names)

    print(f"\n✅ Final merged document created: {output_doc_path}")

def generate_all_cndag_subdocs(
    json_data,
    reference_doc_path,
    output_dir
):
    os.makedirs(output_dir, exist_ok=True)

    jobs = extract_cndag_jobs(json_data)
    results = []

    for job in jobs:
        section = job["section_search"]
        safe_name = section.replace(" ", "_")
        output_path = os.path.join(
            output_dir,
            f"CNDAG_{safe_name}.docx"
        )

        # ✅ step 1: copy original doc
        copy_reference_doc(reference_doc_path, output_path)

        # ✅ step 2: remove everything except this section
        keep_only_section(output_path, section)

        results.append({
            "section": section,
            "file_path": output_path,
            "ds_id": job["ds_id"]
        })

    return results



def create_section_mapping(
    original_json,
    generated_docs,
    generated_content_lookup=None
):
    """
    original_json            -> parsed JSON from base64
    generated_docs           -> output from generate_all_cndag_subdocs()
    generated_content_lookup -> optional dict {sectionName: content}
    """

    mappings = []

    for section in original_json:
        section_name = section.get("sectionName")
        section_number = section.get("sectionNumber")

        for content in section.get("content", []):
            for field in content.get("fields", []):
                jinja_tag = field.get("jinjaTag")

                for condition in field.get("conditions", []):
                    func = condition.get("function", {})
                    for arg in func.get("argList", []):
                        ds = arg.get("dataSource", {})

                        if ds.get("type") == "CNDAG PIPELINE":
                            # find generated subdoc
                            subdoc = next(
                                (
                                    d for d in generated_docs
                                    if d["section"] == ds.get("sectionSearch")
                                ),
                                None
                            )

                            mappings.append({
                                "sectionName": section_name,
                                "sectionNumber": section_number,
                                "jinjaTag": jinja_tag,
                                "contentType": ds.get("type"),
                                "subDocxPath": subdoc["file_path"] if subdoc else None,
                                "generatedContent": (
                                    generated_content_lookup.get(section_name, "")
                                    if generated_content_lookup
                                    else ""
                                )
                            })

    return mappings

if __name__ == "__main__":


    json_data = base64_txt_to_json("json_b64.txt")


    generated_subdocs = generate_all_cndag_subdocs(
        json_data=json_data,
        reference_doc_path="PIPELINE_EXECUTION_UPGRADE 11.docx",
        output_dir="output_docs"
    )

    print("\n✅ Individual CNDAG sub-documents created:")
    for d in generated_subdocs:
        print(f"- {d['section']} -> {d['file_path']}")


    generate_final_cndag_doc(
        json_data=json_data,
        reference_doc_path="PIPELINE_EXECUTION_UPGRADE 11.docx",
        output_doc_path="final_cndag_doc.docx"
    )

    print("\n✅ Final merged CNDAG document created: final_cndag_doc.docx")


    generated_content_lookup = {
        "Troubleshooting Quick Reference": "\n hello "
    }

    mapping_json = create_section_mapping(
        original_json=json_data,
        generated_docs=generated_subdocs,
        generated_content_lookup=generated_content_lookup
    )

    with open("section_mapping.json", "w", encoding="utf-8") as f:
        json.dump(mapping_json, f, indent=4)

    print("\n✅ Section mapping JSON created: section_mapping.json")
