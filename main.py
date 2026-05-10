import base64
import json

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


import os
from docx import Document
from copy import deepcopy
import os

def extract_single_section_to_word(
    reference_doc_path: str,
    section_search: str,
    output_dir: str,
    output_filename: str
):
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_filename)

    ref_doc = Document(reference_doc_path)
    out_doc = Document()

    capture = False

    for block in ref_doc.element.body:
        text_content = ""

        # Extract visible text from block
        for node in block.iter():
            if node.tag.endswith("}t") and node.text:
                text_content += node.text

        text_content = text_content.strip()

        # START capture: subsection title match
        if not capture and section_search.lower() in text_content.lower():
            capture = True

        # STOP capture: next numbered section at same or higher level
        elif capture and text_content and text_content[0].isdigit() and section_search.lower() not in text_content.lower():
            break

        if capture:
            out_doc.element.body.append(deepcopy(block))

    out_doc.save(output_path)
    return output_path


def generate_all_cndag_subdocs(
    json_data,
    reference_doc_path,
    output_dir
):
    jobs = extract_cndag_jobs(json_data)
    results = []

    for job in jobs:
        section = job["section_search"]
        safe_name = section.replace(" ", "_")
        filename = f"CNDAG_{safe_name}.docx"

        path = extract_single_section_to_word(
            reference_doc_path=reference_doc_path,
            section_search=section,
            output_dir=output_dir,
            output_filename=filename
        )

        results.append({
            "section": section,
            "file_path": path,
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