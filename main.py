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

if __name__ == "__main__":
    json_data = base64_txt_to_json("json_b64.txt")

    generated_docs = generate_all_cndag_subdocs(
        json_data=json_data,
        reference_doc_path="PIPELINE_EXECUTION_UPGRADE 11.docx",
        output_dir="output_docs"
    )

    for d in generated_docs:
        print(f"- {d['section']} -> {d['file_path']}")