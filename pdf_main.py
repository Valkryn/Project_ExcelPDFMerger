from pypdf import PdfReader, PdfWriter
import os.path


def create_pdf(Applicant, pdf_path):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    writer.append(reader)
    writer.update_page_form_field_values(
        writer.pages[0],
        {
            'First Name': Applicant['First Name:'],
            'Last Name': Applicant['Last Name:'],
            'Text1': Applicant['Last 4:'],
            'Contact Phone': '718-430-4150',  # this will remain for everyone
            'Email': Applicant['E-mail Address:'].split('@')[0],  # grabs text before @
            'undefined_2': Applicant['E-mail Address:'].split('@')[1],  # grabs text after @
            'Mailing Address': Applicant['Address'],
            'City_2': Applicant['City'],
            'State_2': Applicant['State'],
            'Zipcode_2': Applicant['ZipCode'],
            'Experience in the related field': Applicant['Years:'],
            'Work Address': Applicant['Work_Address'],
            'City': Applicant['Work_City'],
            'State': Applicant['Work_State'],
            'Zipcode': Applicant['Work_Zipcode'],
            'I': f'{Applicant['First Name:']} {Applicant['Last Name:']}'  # Section 3 Declaration
        }, auto_regenerate=True
    )
    while not os.path.isdir('./Merged_files/'):
        os.mkdir('./Merged_files/')
    with open(
            f'Merged_files/{Applicant['First Name:']}_{Applicant['Last Name:']}_{Applicant['Employee Number']}.pdf',
            'wb') as output_stream:
        writer.write(output_stream)
