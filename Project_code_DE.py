import requests
from bs4 import BeautifulSoup
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# Initialize an empty list to store the data
data = []

# Define the URL of the website
url = "https://result.bbiseqta.edu.pk/Result/HSSC"

# Define the range of roll numbers you want to iterate over
start_roll_number = 351001  # Starting roll number
end_roll_number = 367387   # Ending roll number

# Function to add data to the main data list
def add_data_to_list(student_details, part_one_results, part_two_results, total_obtained_marks, remark_status):
    student_data = {
        'Roll No': student_details.get('Roll No', ''),
        'Name': student_details.get('Name', ''),
        'Father Name': student_details.get('Father Name', ''),
        # 'Board Name': board_name,
        # 'Exam Details': exam_details,
    }
    # Add Part One Results as attributes
    for subject_data in part_one_results:
        subject_name = subject_data['subject']
        student_data[f'Part One Result - {subject_name}'] = subject_data['marks']
    # Add Part Two Results as attributes
    for subject_data in part_two_results:
        subject_name = subject_data['subject']
        student_data[f'Part Two Result - {subject_name}'] = subject_data['marks']
    student_data['Total Obtained Marks'] = total_obtained_marks
    student_data['Remarks'] = remark_status
    data.append(student_data)

# Define a function to process a single roll number
def process_roll_number(roll_number):
    # Convert the roll number to a string and pad with leading zeros if needed
    roll_number_str = str(roll_number).zfill(6)

    # Send a GET request to the page to retrieve the verification code
    response = requests.get(url)

    # Check if the request was successful (HTTP status code 200)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'lxml')
        verification_code_element = soup.find('input', {'id': 'a'})
        verification_code = verification_code_element.get('value') if verification_code_element else None

        if verification_code:
            form_data = {
                "Part": "2",
                "Year": "2023",
                "SessCode": "1",
                "RollNo": roll_number_str,
                "vercode": verification_code
            }

            # Send a POST request to submit the form
            response_2 = requests.post(url, data=form_data)

            if response_2.status_code == 200:
                soup_2 = BeautifulSoup(response_2.text, 'lxml')

                # Check if the page contains the board name and exam details
                try:
                    board_name = soup_2.find('h1', class_='head').text.strip()
                    exam_details = soup_2.find('h3', class_='bold').text.strip()
                except AttributeError:
                    # If board name or exam details are not found, return without adding the student data
                    return
                
                # Student details
                student_details = {}
                student_info = soup_2.find_all('p', class_='t-20')
                for info in student_info:
                    try:
                        label = info.find('span', {'style': 'width:150px; float:left; font-weight:bold'}).text.strip()
                        value = info.find('span', class_='bold underline text-capitalize').text.strip()
                        student_details[label] = value
                    except AttributeError:
                        continue

                # Extract results for Part One
                part_one_results = []
                rows = soup_2.find_all('div', class_='row border-sub color-default')
                for row in rows:
                    try:
                        subject = row.find('div', class_='col-8 sub-head sub-title font-wieght-bold bold').text
                        marks = row.find('div', class_='col-4 sub-head text-center sub-title bold').text

                        part_one_results.append({
                            'subject': subject,
                            'marks': marks
                        })
                    except AttributeError:
                        continue

                # Extract results for Part Two
                part_two_results = []
                rows_II = soup_2.find_all('div', class_='row border-sub color-default font-wieght-bold bold')
                for row_II in rows_II:
                    try:
                        subject = row_II.find('div', class_='col-8 sub-head sub-title bold').text
                        marks = row_II.find('div', class_='col-4 sub-head text-center sub-title bold').text

                        part_two_results.append({
                            'subject': subject,
                            'marks': marks
                        })
                    except AttributeError:
                        continue

                try:
                    # Total obtained marks and Remarks
                    sub_titles = soup_2.find_all('div', class_='col-6 sub-head text-center sub-title')

                    # Extract the total marks and remark status
                    total_obtained_marks = sub_titles[0].span.get_text()
                    remark_status = sub_titles[1].span.get_text()
                except AttributeError:
                    total_obtained_marks = ''  # Set as an empty string if not found
                    remark_status = ''  # Set as an empty string if not found

                # Add data to the main data list
                add_data_to_list(student_details, part_one_results, part_two_results, total_obtained_marks, remark_status)
            else:
                print(f"Failed to retrieve the web page for roll number {roll_number}. HTTP status code:", response_2.status_code)
        else:
            print(f"Verification code not found on the page for roll number {roll_number}.")
    else:
        print(f"Failed to retrieve the web page for roll number {roll_number}. HTTP status code:", response.status_code)

# Use ThreadPoolExecutor for parallel processing
with ThreadPoolExecutor(max_workers=10) as executor:
    executor.map(process_roll_number, range(start_roll_number, end_roll_number + 1))

# Create a DataFrame from the list of dictionaries
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
df.to_excel('student_data_PF_1.xlsx', index=False)

print('completed')