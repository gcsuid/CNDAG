import base64
import json
import os
import shutil
from copy import deepcopy
from docx import Document


# ================================
# STEP 1: Base64 → JSON
# ================================
def base64_txt_to_json(input_file, output_file=None):
    with open(input_file, "r") as file:
        base64_str = file.read().strip()

    decoded_str = base64.b64decode(base64_str).decode("utf-8")
    json_data = json.loads(decoded_str)

    if output_file:
        with open(output_file, "w") as outfile:
            json.dump(json_data, outfile, indent=4)

    return json_data


# ================================
# STEP 2: Extract CNDAG jobs
# ================================
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


# ================================
# STEP 3: Copy reference doc
# ================================
def copy_reference_doc(reference_doc_path, output_doc_path):
    shutil.copy(reference_doc_path, output_doc_path)


# ================================
# STEP 4: Keep only required section
# ================================
def keep_only_section(doc_path: str, section_search: str):
    doc = Document(doc_path)
    body = doc.element.body
    blocks = list(body)

    para_map = {p._element: p for p in doc.paragraphs}

    start_idx = None
    start_level = None

    # Find section start
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

    if start_idx is None:
        doc.save(doc_path)
        return

    # Find section end
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

    # Restore only required section
    for block in blocks[start_idx:end_idx]:
        body.append(block)

    doc.save(doc_path)


# ================================
# STEP 5: Generate subsection docs
# ================================
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

        # Copy original document
        copy_reference_doc(reference_doc_path, output_path)

        # Extract only required section
        keep_only_section(output_path, section)

        results.append({
            "section": section,
            "file_path": output_path,
            "ds_id": job["ds_id"]
        })

    return results


# ================================
# ✅ NEW STEP: Merge all subsection docs
# ================================
def merge_subdocs_into_final(subdoc_results, final_doc_path):
    final_doc = Document()

    # Remove default blank paragraph
    if final_doc.paragraphs:
        p = final_doc.paragraphs[0]._element
        p.getparent().remove(p)

    for doc_info in subdoc_results:
        subdoc_path = doc_info["file_path"]
        sub_doc = Document(subdoc_path)

        for element in sub_doc.element.body:
            final_doc.element.body.append(deepcopy(element))

    final_doc.save(final_doc_path)
    print(f"\n✅ Final merged document created: {final_doc_path}")


# ================================
# MAIN
# ================================
if __name__ == "__main__":
    # Step 1: Load JSON
    json_data = base64_txt_to_json("json_b64.txt")

    # Step 2: Generate subsection documents
    generated_docs = generate_all_cndag_subdocs(
        json_data=json_data,
        reference_doc_path="PIPELINE_EXECUTION_UPGRADE 11.docx",
        output_dir="output_docs"
    )

    print("\n✅ Sub-documents created:")
    for d in generated_docs:
        print(f"- {d['section']} -> {d['file_path']}")

    # Step 3: Merge into final document
    merge_subdocs_into_final(
        subdoc_results=generated_docs,
        final_doc_path="final_cndag_doc.docx"
    )