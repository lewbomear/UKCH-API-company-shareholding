import docx
import csv
import os
import requests
from print_shareholder_info import print_shareholder_info
from dotenv import load_dotenv


def generate_relevant_individual_info():
    """
    Generates information for an individual associate's company data
    """
    load_dotenv()

    document = docx.Document()

    api_key = os.environ.get("COMPANY_HOUSE_API_KEY")
    officer_name = os.environ.get("OFFICER_NAME")
    officer_dob = os.environ.get("OFFICER_DOB")

    url = f'https://api.company-information.service.gov.uk/search/officers?q="{officer_name}"'

    headers = {"Authorization": api_key}

    # Add a title to the document
    document.add_heading(f"Associated companies for: {officer_name}", level=1)

    # Load the SIC codes and descriptions from the CSV file
    with open(
        os.path.join(
            os.path.dirname(os.path.realpath(__file__)),
            "SIC07_CH_condensed_list_en.csv",
        ),
        newline="",
    ) as csvfile:
        sic_codes_reader = csv.reader(csvfile, delimiter=",", quotechar='"')
        sic_mapping = {row[0]: row[1] for row in sic_codes_reader}

    # Make the API request
    response = requests.get(url, headers=headers)
    data = response.json()

    # only keep officers with exact name match
    exact_name_matches = []
    for item in data["items"]:
        if item["title"].lower() == officer_name.lower():
            exact_name_matches.append(item)

    exact_name_dob_matches = []
    for item in data["items"]:
        if (
            item["title"].lower() == officer_name.lower()
            and isinstance(item.get("date_of_birth"), dict)
            and str(item["date_of_birth"]["year"])
            + "-"
            + str(item["date_of_birth"]["month"]).zfill(2)
            == officer_dob[:7]
        ):
            exact_name_dob_matches.append(item)

    # Loop through the exact name matches and make a request to each URL in the links dictionary
    for match in exact_name_dob_matches:
        officer_url = (
            f"https://api.company-information.service.gov.uk{match['links']['self']}"
        )
        officer_response = requests.get(officer_url, headers=headers)
        officer_data = officer_response.json()

        # Loop through the officer's appointments and print the company name, number, and nature of business
        for appointment in officer_data["items"]:
            company_name = appointment["appointed_to"]["company_name"]
            company_number = appointment["appointed_to"]["company_number"]
            appointed_on = appointment["appointed_on"]
            officer_role = appointment["officer_role"]
            psc_name = ""

            # Get the company profile URL
            company_profile_url = f"https://api.company-information.service.gov.uk/company/{company_number}"
            company_profile_response = requests.get(
                company_profile_url, headers=headers
            )
            company_profile_data = company_profile_response.json()

            # Get the SIC code and look up the activity in the mapping dictionary
            sic_code = company_profile_data.get("sic_codes", ["N/A"])[0]
            activity = sic_mapping.get(sic_code, "Unknown")

            # Get the company's persons with significant control data
            company_psc_url = f"https://api.company-information.service.gov.uk/company/{company_number}/persons-with-significant-control"
            company_psc_response = requests.get(company_psc_url, headers=headers)
            company_psc_data = company_psc_response.json()


            # Check if the 'items' key is present in the company_psc_data dictionary
            if "items" not in company_psc_data:
                if "resigned_on" in appointment:
                    resigned_on = appointment["resigned_on"]
                    document.add_paragraph(
                        f"{officer_name} was appointed a {officer_role} of {company_name} ({company_number}) on {appointed_on} and resigned on {resigned_on}. The nature of business is {activity}. The company has no persons with significant control."
                    )
                else:
                    document.add_paragraph(
                        f"{officer_name} has been serving as {officer_role} of {company_name} ({company_number}) since {appointed_on}. The nature of business is {activity}. The company has no persons with significant control."
                    )
                continue

            # Loop through the company's persons with significant control and print the name
            psc_names = []
            for psc in company_psc_data["items"]:
                if psc["name"]:
                    psc_name = psc["name"]
                """else:
                    if psc["title"]:
                        psc_name = psc["title"]
                    if psc["forename"]:
                        psc_name += " "
                        psc_name += psc["forename"]
                    if psc["middle_name"]:
                        psc_name += " "
                        psc_name += psc["middle_name"]
                    if psc["surname"]:
                        psc_name += " "
                        psc_name += psc["surname"]"""
                psc_names.append(psc_name)

            print(psc_names)

            if "resigned_on" in appointment:
                resigned_on = appointment["resigned_on"]
                if len(psc_names) == 1:
                    for psc_name in psc_names:
                        psc_statement = f"The company has a person with significant control named {psc_name}."
                    document.add_paragraph(
                        f"{officer_name} has been serving as {officer_role} of {company_name} ({company_number}) since {appointed_on}. The nature of business is {activity}. {psc_statement}"
                    )
                elif len(psc_names) > 1:
                    last_name = psc_names.pop()
                    full_list = ", ".join(psc_names)
                    psc_statement = f"The company has a persons with significant control named {full_list} and {last_name}."
                    document.add_paragraph(
                        f"{officer_name} has been serving as {officer_role} of {company_name} ({company_number}) since {appointed_on}. The nature of business is {activity}. {psc_statement}"
                    )
                else:
                    psc_statement = (
                        "The company has no persons with significant control"
                    )
                    document.add_paragraph(
                        f"{officer_name} has been serving as {officer_role} of {company_name} ({company_number}) since {appointed_on}. The nature of business is {activity}. {psc_statement}"
                    )
            else:
                if len(psc_names) == 1:
                    for psc_name in psc_names:
                        psc_statement = f"The company has a person with significant control named {psc_name}."
                    document.add_paragraph(
                        f"{officer_name} has been serving as {officer_role} of {company_name} ({company_number}) since {appointed_on}. The nature of business is {activity}. {psc_statement}"
                    )
                elif len(psc_names) > 1:
                    last_name = psc_names.pop()
                    full_list = ", ".join(psc_names)
                    psc_statement = f"The company has a persons with significant control named {full_list} and {last_name}."
                    document.add_paragraph(
                        f"{officer_name} has been serving as {officer_role} of {company_name} ({company_number}) since {appointed_on}. The nature of business is {activity}. {psc_statement}"
                    )
                else:
                    psc_statement = (
                        "The company has no persons with significant control"
                    )
                document.add_paragraph(
                    f"{officer_name} has been serving as a {officer_role} of {company_name} ({company_number}) since {appointed_on}. The nature of business is {activity}. {psc_statement}"
                )

    # Save the document as a Word file
    document.save(f"Associated companies for {officer_name}.docx")
    document_text = docx.Document(f"Associated companies for {officer_name}.docx")
    print(document_text)


if __name__ == "__main__":
    generate_relevant_individual_info()
