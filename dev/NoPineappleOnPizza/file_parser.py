from bs4 import BeautifulSoup
import sys

def extract_students(xml_data):
    soup = BeautifulSoup(xml_data, 'html.parser')
    students = soup.find_all('student')

    for student in students:
        raw_text = student.get_text(separator='\n').strip().split('\n')
        raw_text = [line.strip() for line in raw_text if line.strip()]

        # Try to find the <Name> tag
        name_tag = student.find('name')
        if name_tag:
            name = name_tag.text.strip()
        else:
            # Use the first line that's not a number or "Lab"
            name = next((line for line in raw_text if not line.isdigit() and "Lab" not in line and len(line.split()) >= 2), "Unknown")

        # Try to find the <LabVisit> tag
        lab_tag = student.find('labvisit')
        if lab_tag:
            lab = lab_tag.text.strip()
        else:
            # Find the first line that has the word "Lab" in it
            lab = next((line for line in raw_text if "Lab" in line), "Unknown")

        print(f"Name: {name}")
        print(f"Lab: {lab}")
        print("-" * 30)

# Run the script from command line with file path
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 file_parser.py <file_path>")
        sys.exit(1)

    file_path = sys.argv[1]
    with open(file_path, 'r') as f:
        xml_content = f.read()

    extract_students(xml_content)

