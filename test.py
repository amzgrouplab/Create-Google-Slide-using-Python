import os
import json
import csv
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Load the service account credentials
credentials_file = 'credentials.json'
credentials = service_account.Credentials.from_service_account_file(credentials_file)
scopes = ['https://www.googleapis.com/auth/presentations']

# Build the Google Slides API service
service = build('slides', 'v1', credentials=credentials)

def add_slide(presentation_id):
    # Call the presentations().get() method to retrieve the presentation
    presentation = service.presentations().get(presentationId=presentation_id).execute()

    # Retrieve the first slide from the presentation
    slide_id = presentation['slides'][0]['objectId']
    # Load the specification data from the CSV file
    data = []
    with open('slide-specifications.csv', 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            data.append(row)
    print(data)
    num_rows = len(data)
    print(num_rows)
    #X, Y coordinates for the start of the title block
    title_x = float(data[0][0])
    title_y = float(data[0][1])
    
    #X, Y coordinates for the start of the table
    table_x = float(data[1][0])
    table_y = float(data[1][1])
    
    #widths of the columns
    column_widths = [x for x in data[2] if isinstance(x, int) or (isinstance(x, str) and x.isdigit())]
    column_widths = [float(x) for x in column_widths]
    #height of the row
    row_height = float(data[3][0])
    
    #This is the text justification side
    text_justifications = data[4]
    
    print(title_x, title_y, table_x, table_y, column_widths, row_height, text_justifications)
    # Load the slide data from the CSV file
    data = []
    with open('Slide-data.csv', 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            data.append(row)
            
    title = data[0][0]
    data = data[1:]

    print('Title:', title)
    print('Data:', data)
    
    # Create a title
    requests = [
        {
            'createShape': {
                'objectId': 'title_object_id',
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'width': {
                            'magnitude': 200,
                            'unit': 'PT'
                        },
                        'height': {
                            'magnitude': 30,
                            'unit': 'PT'
                        }
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': title_x*72,  # Left position in PT (1 inch)
                        'translateY': title_y*-36,  # Top position in PT (0.5 inch)
                        'unit': 'PT'
                    }
                }
            }
        },
        {
            'insertText': {
                'objectId': 'title_object_id',
                'text': title,
                'insertionIndex': 0
            }
        },
        {
            'updateTextStyle': {
                'objectId': 'title_object_id',
                'textRange': {
                    'type': 'ALL'
                },
                'style': {
                    'fontFamily': 'Montserrat',
                    'fontSize': {
                        'magnitude': 8,
                        'unit': 'PT'
                    },
                    'foregroundColor': {
                        'opaqueColor': {
                            'rgbColor': {
                                'red': 1.0,
                                'green': 1.0,
                                'blue': 1.0
                            }
                        }
                    }
                },
                'fields': 'foregroundColor,fontFamily,fontSize'
            }
        }
    ]

    service.presentations().batchUpdate(
        presentationId=presentation_id, body={'requests': requests}).execute()

    # Determine the number of rows and columns
    num_rows = len(data)
    num_cols = len(data[0]) if data else 0
    
    # Create a table
    requests = [
        {
            'createTable': {
                'objectId': 'table_object_id',
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'width': {
                            'magnitude': 400,
                            'unit': 'PT'
                        },
                        'height': {
                            'magnitude': 200,
                            'unit': 'PT'
                        }
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': table_x*72,
                        'translateY': table_y*-72,
                        'unit': 'PT'
                    }
                },
                'rows': num_rows,  # number of rows in the table, including header row
                'columns': num_cols,  # number of columns in the table
            }
        },
        {
            'updateTableCellProperties': {
                'objectId': 'table_object_id',
                'tableRange': {
                    'location': {
                        'rowIndex': 0,
                        'columnIndex': 0
                    },
                    'rowSpan': 4,
                    'columnSpan': 4
                },
                'tableCellProperties': {
                    'tableCellBackgroundFill': {
                        'propertyState': 'NOT_RENDERED'
                    }
                },
                'fields': 'tableCellBackgroundFill'
            }
        },
    ]


    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body={'requests': requests}).execute()

    table_id = response['replies'][0]['createTable']['objectId']
    
    
    # Populate the table with data
    requests = []
    for row in range(len(data)):
        for col in range(len(data[row])):
            text = data[row][col]
            if text:
                requests.append({
                    'insertText': {
                        'objectId': table_id,
                        'cellLocation': {
                            'rowIndex': row,
                            'columnIndex': col
                        },
                        'text': data[row][col]
                    }
                })
                requests.append({
                    'updateTextStyle': {
                        'objectId': table_id,
                        'cellLocation': {
                            'rowIndex': row,
                            'columnIndex': col
                        },
                        'style': {
                            'foregroundColor': {
                                'opaqueColor': {
                                    'rgbColor': {
                                        'red': 1.0,
                                        'green': 1.0,
                                        'blue': 1.0
                                    }
                                }
                            },
                            'fontSize': {
                                'magnitude': 9,
                                'unit': 'PT'
                            },
                            'fontFamily': 'Montserrat',
                        },
                        'textRange': {
                            'type': 'ALL'
                        },
                        'fields': 'foregroundColor,fontSize,fontFamily'
                    }
                })
    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body={'requests': requests}).execute()
    # Set the column widths (in PT) - Modify these values as per your requirements
    column_widths = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]

    # Update column widths
    requests = []
    for col in range(num_cols):
        column_width = column_widths[col]
        requests.append({
            'updateTableColumnProperties': {
                'objectId': table_id,
                'columnIndices': [col],
                'tableColumnProperties': {
                    'columnWidth': {
                        'magnitude': column_width * 58,
                        'unit': 'PT'
                    }
                },
                'fields': 'columnWidth'
            }
        })
    requests.append({
        'updateTableRowProperties': {
            'objectId': table_id,
            'rowIndices': [0],  # Index of the first row
            'tableRowProperties': {
                'minRowHeight': {
                    'magnitude': 30,
                    'unit': 'PT'
                }
            },
            'fields': 'minRowHeight'
        }
    })
    for row in range(1, num_rows):
        requests.append({
            'updateTableRowProperties': {
                'objectId': table_id,
                'rowIndices': [row],  # Index of the current row
                'tableRowProperties': {
                    'minRowHeight': {
                        'magnitude': 10,
                        'unit': 'PT'
                    }
                },
                'fields': 'minRowHeight'
            }
        })
    # Apply row color to the table
    for row in range(num_rows):
            if row != 0:
                if row % 2 == 0:
                    # Apply color to even-indexed rows starting from the second row
                    row_color = {
                        'solidFill': {
                            'color': {
                                'rgbColor': {
                                    'red': 0.0,
                                    'green': 0.7,
                                    'blue': 0.8
                                }
                            }
                        }
                    }
                else:
                    # Apply color to odd-indexed rows and the first row
                    row_color = {
                        'solidFill': {
                            'color': {
                                'rgbColor': {
                                    'red': 0,
                                    'green': 0.3,
                                    'blue': 0.8
                                }
                            }
                        }
                    }

                requests.append({
                    'updateTableCellProperties': {
                        'objectId': table_id,
                        'tableRange': {
                            'location': {
                                'rowIndex': row,
                                'columnIndex': 0
                            },
                            'rowSpan': 1,
                            'columnSpan': num_cols
                        },
                        'tableCellProperties': {
                            'tableCellBackgroundFill': row_color
                        },
                        'fields': 'tableCellBackgroundFill'
                    }
                })

    # Execute the batchUpdate request to update column widths
    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body={'requests': requests}).execute()
    
  

    print(f"New slide created with ID: {table_id}")
if __name__ == '__main__':
    presentation_id = '1FESkY8xzKRCLKfioTW6NKzJRjgG_yMggJAuYi-nyqxU'  # Replace with your presentation ID
    add_slide(presentation_id)
