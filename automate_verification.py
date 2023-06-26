import os
import docx
import datetime

from docx.api import Document
from fpdf import FPDF

VERIFICATION_TEST_PATH = r"C:\verification_tests"
OUTPUT_PATH = ""
TESTER_SIGNATURE = r"C:\Users\admin\Desktop\signature.JPEG"


def pre_conditions(file_path):
    doc = docx.Document(file_path)
    paragraphs = []
    for paragraph in doc.paragraphs:
        paragraphs.append(paragraph.text)


    for line in paragraphs:
        if "Pre conditions" in line:
            return line

    return "Couldn't find the Pre conditions instructions"


def read_table():
    keys = None
    for i, row in enumerate(table.rows):
        text = []
        for cell in row.cells:
            text.append(cell.text)

        # Establish the mapping based on the first row
        # headers; these will become the keys of our dictionary
        if i == 0:
            keys = tuple(text)
            continue

        # Construct a dictionary for this row, mapping
        # keys to values for this row
        row_data = dict(zip(keys, text))
        data.append(row_data)


def generate_pdf(docx_path, pdf_path):
    # Open the DOCX file
    doc = Document(docx_path)

    # Create a PDF object
    pdf = FPDF('P', 'mm', 'A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    first_time = True
    write_table = False

    # Iterate over the paragraphs in the document
    for paragraph in doc.paragraphs:
        if first_time:
            pdf.set_font("Arial", size=16, style="B")
            pdf.cell(0, 10, txt=paragraph.text, ln=1, align="C")
            first_time = False
            continue

        pdf.set_font("Arial", size=12)

        if "Tester" in paragraph.text:
            x = pdf.get_x() + 150
            y = pdf.get_y() - 5
            pdf.cell(0, 10, txt="Tester: " + tester_name + "\t" * 75 + "Signature: ", ln=1)

            pdf.image(TESTER_SIGNATURE, x=x, y=y, w=30, h=20)  # Add the signature image
            continue

        if "Date" in paragraph.text:
            pdf.cell(0, 10, txt="Date: " + datetime.datetime.now().strftime("%Y-%m-%d"), ln=1)
            continue

        if "version" in paragraph.text:
            pdf.cell(0, 10, txt="SW version \ build number: " + version_number, ln=1)
            continue

        if "Summary" in paragraph.text:
            pdf.cell(0, 10, txt="Summary, conclusion and recommendations: " + summary, ln=1)
            continue

        if write_table:
            pdf.set_font("Arial", size=10)
            first_row = True
            test_info = True
            result_info = False
            result = False
            index = 1

            for row_data in data:
                for head_line, to_print in row_data.items():
                    # skip first row
                    if first_row:
                        first_row = False
                        continue

                    # first col
                    if test_info:
                        if to_print == "":
                            first_row = True
                            test_info = True
                            continue

                        pdf.cell(0, 7, txt=str(index) + ". Test: " + to_print, ln=1)
                        test_info = False
                        result_info = True

                    # second col
                    elif result_info:
                        pdf.cell(0, 7, txt="      Expected_Results: " + to_print, ln=1)
                        result_info = False
                        result = True


                    elif result:
                        pdf.cell(40, 7, txt="      Result:", ln=0)

                        # Set the text color based on the result
                        if test_results[index - 1] == "v":
                            pdf.set_text_color(0, 128, 0)  # Set text color to green for "Pass"
                            pdf.cell(35, 7, txt="Pass", ln=1, align='L')
                        else:
                            pdf.set_text_color(255, 0, 0)  # Set text color to red for "Fail"
                            pdf.cell(35, 7, txt="Fail", ln=1, align='L')

                        # Reset the text color to black
                        pdf.set_text_color(0, 0, 0)
                        result = False

                    # fourth col
                    else:
                        if comment_results[index - 1] != "":
                            pdf.cell(0, 7, txt="      Comments: " + comment_results[index - 1],
                                     ln=1)
                        first_row = True
                        test_info = True
                        index += 1

            pdf.ln()
            write_table = False

        if "Pre conditions" in paragraph.text:
            write_table = True

        pdf.cell(0, 10, txt=paragraph.text, ln=1)

    # Save the PDF file
    with open(pdf_path, 'wb') as file:
        file.write(pdf.output(dest='S').encode('latin-1'))






for test_type in os.listdir(VERIFICATION_TEST_PATH):
    print(test_type.upper())
    version_number = input("please insert the version: ")

    for test_file in os.listdir(os.path.join(VERIFICATION_TEST_PATH, test_type)):
        pdf_exist = False
        for f in os.listdir(os.path.join(VERIFICATION_TEST_PATH, test_type)):
            if f.endswith('pdf'):
                pdf_exist = True
                break

        if pdf_exist:
            cont = input('this test case has already been tested, would you like to skip? y/n ')
            while cont != 'y' and cont != 'n':
                cont = input('Illegal input \nthis test case has already been tested, would you like to skip? y/n ')
            if cont == 'y':
                continue

        tester_name = input("please insert tester name: ")
        test_path = os.path.join(VERIFICATION_TEST_PATH, test_type, test_file)

        date_string = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        pdf_name = test_file.replace(".docx", f"_{date_string}.pdf")
        OUTPUT_PATH = os.path.join(VERIFICATION_TEST_PATH, test_type, pdf_name)

        print(f"\nFile name: {test_file}")
        print(pre_conditions(test_path) + "\n")

        # save input
        test_results = []
        comment_results = []
        idx = 0

        # read table
        document = Document(test_path)
        table = document.tables[0]
        data = []
        read_table()

        first_row = True
        test_info = True
        result_info = False
        result = False

        for idx, row_data in enumerate(data):
            if idx == len(data) - 3:
                break
            for head_line, to_print in row_data.items():
                # skip first row
                if first_row:
                    first_row = False
                    continue

                print(f"{head_line}:", end=" ")

                # first col
                if test_info:
                    if to_print == "":
                        first_row = True
                        test_info = True
                        continue

                    print(f"{to_print}")
                    test_info = False
                    result_info = True

                # second col
                elif result_info:
                    print(f"{to_print}")
                    result_info = False
                    result = True

                # third col
                elif result:
                    while True:
                        res = input("v/x? ").lower()
                        if res in ['v', 'x']:
                            test_results.append(res)
                            result = False
                            break
                        else:
                            print("Choose 'v' or 'x'")

                # fourth col
                else:
                    print("(press enter to continue)")
                    comment = input().lower()
                    comment_results.append(comment)
                    first_row = True
                    test_info = True
                    idx += 1

        summary = input("Summary, conclusion and recommendations: (press enter to continue)")
        generate_pdf(test_path, OUTPUT_PATH)
