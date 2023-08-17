from __future__ import print_function
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
]
# Environment variable for Sources Folder ID
SOURCES_FOLDER_ID = os.environ['SOURCES_FOLDER_ID']


def start_api():
    """Generates credentials from credentials.json file."""

    creds = None
    """The file token.json stores the user's access and refresh tokens, and is
    created automatically when the authorization flow completes for the first
    time."""
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def move_doc_to_sources(file_id, folder_id, creds):
    """Moves the doc created in the Root folder to the 'Sources' folder."""

    try:
        # call drive api client
        service = build('drive', 'v3', credentials=creds)

        # Retrieve the existing parents to remove
        file = service.files().get(fileId=file_id, fields='parents').execute()
        previous_parents = ','.join(file.get('parents'))

        # Move the file to the new folder
        service.files().update(fileId=file_id, addParents=folder_id, removeParents=previous_parents,
                               fields='id, parents').execute()

        print('Doc moved to the Sources folder')

    except HttpError as error:
        print(f'An error occurred while moving the doc to sources folder: {error}')
        return None


def create_doc(data, creds):
    """Creates a blank doc in Root folder."""

    try:
        service = build('docs', 'v1', credentials=creds)

        doc_title = data['properties']['title']
        body = {
            'title': doc_title
        }

        doc = service.documents().create(body=body).execute()
        print(f'Created document with the title: {doc.get("title")}')
        doc_id = doc.get('documentId')
        return doc_id

    except HttpError as error:
        print(f'An error occurred while creating the doc: {error}')
        return None


def process_data(data):
    """Reads the data provided by read_scripts() and identifies the section headings, the links under
    those headings and the title of the script."""

    headings = []
    links = []
    row_data = data['sheets'][0]['data'][0]['rowData']
    row_num = 0
    title = row_data[0]['values'][0]['userEnteredValue']['stringValue']

    try:
        # Get the links and headings with their respective row numbers
        for row in row_data:
            text_is_strikethrough = row['values'][0]['effectiveFormat']['textFormat']['strikethrough']
            text_is_bold = row['values'][0]['effectiveFormat']['textFormat']['bold']

            # Check if link display is present
            try:
                text_type = row['values'][0]['effectiveFormat']['hyperlinkDisplayType']
                # Skip the lines that have been strikethrough
                if text_is_strikethrough is True:
                    continue

                if text_type == 'LINKED':
                    # If present, append the links and its row num to the list in each cell
                    for text_format in row['values'][0]['textFormatRuns']:
                        try:
                            link = text_format['format']['link']['uri']
                            links.append([link, row_num])
                        # If link not present, continue
                        except (KeyError, IndexError):
                            continue

            # If link display is not present (Text is plain text), ignore it
            except (KeyError, IndexError):
                pass

            # Check if the text is Bold (Heading)
            else:
                text_align = row['values'][0]['effectiveFormat']['horizontalAlignment']
                if text_is_bold is True and text_align == "CENTER":
                    heading = row['values'][0]['userEnteredValue']['stringValue']
                    headings.append([heading, row_num])

            finally:
                row_num += 1

    except HttpError as error:
        print(f'An error occurred while processing the data: {error}')
        return None

    finally:
        return headings, links, title


def add_data_to_doc(headings, links, title, doc_id, creds):
    """Adds the data to the created doc in proper formatting."""

    # Make a dictionary of headings as keys and links as value lists under the proper heading
    links_under_headings = {}
    sub_links = []

    # Add the last line to the nested list. So the heading list now contains [Heading, first row, last row]
    for i in range(len(headings) - 1):
        current_heading = headings[i]
        next_heading = headings[i + 1]
        current_heading.append(next_heading[1])

    # For the last heading's last line, add 1000
    last_heading = headings[-1]
    last_heading.append(1000)
    for heading in headings:
        # Ignore the first heading as it is the title
        if heading[1] == 0:
            continue
        else:
            for link in links:
                if heading[1] < link[1] < heading[2]:
                    sub_links.append(link[0])

        # Only append to the dictionary if the links list is not empty
        if not sub_links:
            continue
        else:
            links_under_headings[heading[0]] = list(set(sub_links))
            sub_links = []

    try:
        service = build('docs', 'v1', credentials=creds)

        # Add title to the Doc
        title_request = [
            {
                'insertText': {
                    'location': {
                        'index': 1,
                    },
                    'text': f'{title}\n\n\n'
                }
            },
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': 1,
                        'endIndex': len(title)
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'TITLE',
                        'spaceAbove': {
                            'magnitude': 10.0,
                            'unit': 'PT'
                        },
                        'spaceBelow': {
                            'magnitude': 10.0,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'namedStyleType,spaceAbove,spaceBelow'
                }
            }
        ]

        # Execute the request to add title to the doc
        service.documents().batchUpdate(documentId=doc_id, body={'requests': title_request}).execute()

        requests = []
        current_index = len(title) + 3

        # Add the headings in Bold and links under the heading in bulleted format
        for heading, links in links_under_headings.items():
            requests.append({
                'insertText': {
                    'location': {
                        'index': current_index,
                    },
                    'text': f'{heading}\n\n'
                }
            })
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(heading)
                    },
                    'textStyle': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            })

            # Move to the next line
            current_index += len(heading) + 1
            bullet_index_start = current_index

            # Add the links under the headings
            for link in links:
                requests.append({
                    'insertText': {
                        'location': {
                            'index': current_index,
                        },
                        'text': f'{link}\n'
                    }
                })
                requests.append({
                    'updateTextStyle': {
                        "range": {
                            'startIndex': current_index,
                            'endIndex': current_index + len(link)
                        },
                        'textStyle': {
                            'link': {
                                'url': link
                            }
                        },
                        'fields': 'link'
                    }
                })
                current_index += len(link) + 1

            # Make the links bulleted
            requests.append(
                {
                    'createParagraphBullets': {
                        'range': {
                            'startIndex': bullet_index_start,
                            'endIndex': current_index
                        },
                        'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE',
                    }
                })

            # Leave a blank line between headings
            current_index += 1

        # Add the headings and links to the doc
        service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
        print('Added the data to the doc successfully!')

    except HttpError as error:
        print(f'An error occurred while adding the data to the doc: {error}')
        return None


def read_script(sheet_id, sheet_name, creds):
    """Open the script and extract all the data from Column A."""

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        # Get the entire column A data from the given Sheet
        data = sheet.get(spreadsheetId=sheet_id,
                         includeGridData=True,
                         ranges=f'{sheet_name}!A:A').execute()

        return data

    except HttpError as error:
        print(f'An error occurred while reading the script from the sheet: {error}')
        return None


def main():
    # Verify the credentials
    creds = start_api()

    # Ask for the Sheet ID and name
    sheet_name = input('Enter the full name (i.e.version) of the Sheet: ')
    sheet_id = input('Enter the Sheet ID: ')

    # Read the script and extract the data
    data = read_script(sheet_id, sheet_name, creds)
    headings, links, title = process_data(data)
    title += ' Sources'

    # Create the doc, add the data to it and move it to the Sources folder
    doc_id = create_doc(data, creds)
    add_data_to_doc(headings, links, title, doc_id, creds)
    move_doc_to_sources(doc_id, SOURCES_FOLDER_ID, creds)


if __name__ == '__main__':
    main()
