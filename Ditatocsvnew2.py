import os
import csv
import json
from bs4 import BeautifulSoup


def extract_premise_and_requirement(file_path):
    """
    Extracts Premise and Requirement sections from a DITA file using BeautifulSoup.
    """
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            xml_content = file.read()

        soup = BeautifulSoup(xml_content, "xml")
        premise = ""
        requirement = ""

        for entry in soup.find_all("entry"):
            paragraphs = entry.find_all("p")
            for p in paragraphs:
                text = p.get_text(strip=True)
                if text.startswith("Premise:"):
                    premise = text.replace("Premise:", "").strip()
                elif text.startswith("Requirement:"):
                    requirement += text.replace("Requirement:", "").strip() + "\n"

        return f"Premise:\n{premise}\nRequirement:\n{requirement.strip()}"
    except Exception as e:
        print(f"Error extracting Premise and Requirement from {file_path}: {e}")
        return ""


def extract_xd_note(file_path):
    """
    Extract xd Notes from the provided _testcase.dita file using BeautifulSoup.
    """
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            xml_content = file.read()

        soup = BeautifulSoup(xml_content, "xml")
        xd_notes = []
        seen_notes = set()

        for entry in soup.find_all("entry"):
            for p in entry.find_all("p"):
                if p.get_text(strip=True).lower() == "xd note":
                    # Collect all sibling notes
                    for sibling in entry.find_all(["li", "p"]):
                        note_text = sibling.get_text(strip=True)
                        if note_text and note_text not in seen_notes:
                            xd_notes.append(note_text)
                            seen_notes.add(note_text)

        return "\n".join(xd_notes)
    except Exception as e:
        print(f"Error extracting xd Notes from {file_path}: {e}")
        return ""


def extract_section_content(soup, section_name):
    """
    Extracts specific sections (e.g., Expected Results, Procedures, Notes) using BeautifulSoup.
    """
    try:
        content = []
        for dlentry in soup.find_all("dlentry"):
            dt = dlentry.find("dt")
            dd = dlentry.find("dd")

            if dt and dd and section_name.lower() in dt.get_text(strip=True).lower():
                # Extract all text within the <dd> tag
                content.append(dd.get_text(strip=True))

        return "\n".join(content)
    except Exception as e:
        print(f"Error extracting section '{section_name}': {e}")
        return ""


def parse_testcase_dita(file_path):
    """
    Parses a _testcase.dita file for xd Note, Expected Results, Procedures, and Notes.
    """
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            xml_content = file.read()

        soup = BeautifulSoup(xml_content, "xml")
        title_tag = soup.find("title")
        title = title_tag.get_text(strip=True) if title_tag else "Unknown Title"

        xd_note = extract_xd_note(file_path)

        custom_fields = []
        for dlentry in soup.find_all("dlentry"):
            dt = dlentry.find("dt")
            if dt and dt.get_text(strip=True).startswith("R"):
                entry_id = dt.get_text(strip=True)
                expected_results = extract_section_content(soup, "Expected Results")
                procedures = extract_section_content(soup, "Procedures")
                notes = extract_section_content(soup, "Notes")

                custom_fields.append({
                    "fields": {
                        "Action": f"*{entry_id}*\n{procedures}",
                        "Data": f"*Notes:*\n{notes}",
                        "Expected Result": expected_results
                    }
                })

        return title, xd_note, custom_fields
    except Exception as e:
        print(f"Error parsing {file_path}: {e}")
        return None, None, None


def parse_dita(file_path):
    """
    Parses a `.dita` file for the general description, Premise, and Requirements.
    """
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            xml_content = file.read()

        soup = BeautifulSoup(xml_content, "xml")
        title_tag = soup.find("title")
        title = title_tag.get_text(strip=True) if title_tag else "Unknown Title"

        description = []
        premise = ""
        requirements = ""

        for dlentry in soup.find_all("dlentry"):
            dt = dlentry.find("dt")
            dd = dlentry.find("dd")

            if dt and dd:
                header = dt.get_text(strip=True).lower()
                body = dd.get_text(strip=True)

                if header == "premise":
                    premise = body
                elif header == "requirements":
                    requirements = body
                else:
                    description.append(f"{dt.get_text(strip=True)}:\n{body}")

        return title, "\n\n".join(description), premise, requirements
    except Exception as e:
        print(f"Error parsing {file_path}: {e}")
        return None, None, None, None


def process_folder(input_folder, output_csv):
    """
    Processes all matching .dita file pairs and outputs them in a structured CSV format.
    """
    try:
        files = [f for f in os.listdir(input_folder) if f.endswith(".dita")]
        base_files = {f.rsplit("_testcase.dita", 1)[0]: f for f in files if f.endswith("_testcase.dita")}
        matched_files = [
            (f"{base_name}.dita", testcase_file)
            for base_name, testcase_file in base_files.items()
            if f"{base_name}.dita" in files
        ]

        data = []
        for base_file, testcase_file in matched_files:
            base_file_path = os.path.join(input_folder, base_file)
            testcase_file_path = os.path.join(input_folder, testcase_file)

            custom_field_premise = extract_premise_and_requirement(base_file_path)

            base_title, base_description, premise, _ = parse_dita(base_file_path)
            testcase_title, xd_note, custom_fields = parse_testcase_dita(testcase_file_path)

            if base_title and testcase_title:
                data.append([
                    base_title,
                    base_description,
                    custom_field_premise,
                    json.dumps(custom_fields),
                    xd_note
                ])

        with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([
                "Summary",
                "Description",
                "Custom field (Premise)",
                "Custom Field Manual Test Steps",
                "Custom Field (xd Notes)"
            ])
            writer.writerows(data)

        print(f"CSV file created successfully: {output_csv}")
    except Exception as e:
        print(f"Error processing folder: {e}")


# Example usage
input_folder = "TRC"  # Replace with the folder containing your .dita files
output_csv = "structured_output.csv"
process_folder(input_folder, output_csv)
