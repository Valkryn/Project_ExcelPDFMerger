from pypdf import PdfReader, PdfWriter
import os.path


def create_pdf(applicant, pdf_path):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    writer.append(reader)
    writer.update_page_form_field_values(
        writer.pages[0],
        {
            'First Name': applicant['First Name:'],
            'Last Name': applicant['Last Name:'],
            'Text1': applicant['Last 4:'],
            'Contact Phone': '718-430-4150',  # this will remain for everyone
            'Email': applicant['E-mail Address:'].split('@')[0],  # grabs text before @
            'undefined_2': applicant['E-mail Address:'].split('@')[1],  # grabs text after @
            'Mailing Address': applicant['Address'],
            'City_2': applicant['City'],
            'State_2': applicant['State'],
            'Zipcode_2': applicant['ZipCode'],
            'Experience in the related field': applicant['Years:'],
            'Work Address': applicant['Work_Address'],
            'City': applicant['Work_City'],
            'State': applicant['Work_State'],
            'Zipcode': applicant['Work_Zipcode'],
            'I': f'{applicant["First Name:"]} {applicant["Last Name:"]}'  # Section 3 Declaration
        }, auto_regenerate=True
    )
    while not os.path.isdir('./Merged_files/'):
        os.mkdir('./Merged_files/')
    with open(
            f'Merged_files/{applicant["First Name:"]}_{applicant["Last Name:"]}_{applicant["Employee Number"]}.pdf',
            'wb') as output_stream:
        writer.write(output_stream)
