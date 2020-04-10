import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import yaml
import clize

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/presentations', "https://www.googleapis.com/auth/drive", 
        "https://www.googleapis.com/auth/drive.file"]

# Yellow arrow image used in Real Python Tables of Contents
IMG_URL = 'https://lh6.googleusercontent.com/kOawokv7LMUFgMDLmuL543c91FMd3ZRLkVK4E-9SwufuIv_RjyOJ_2hFBa2t_OTTRwXrvo-BHrwNph-EhpDQhCrh3LE0Fc40h98_VyILieAe3dfwpD7iu_GQwTs0QuztrbRnPVEoaEmq4k94b1Q'

# def get_creds('token.pickle'='token.pickle', creds_path='credentials/credentials.json'):
    
class ConfigurationError(Exception):
    def __init__(self, message):
        self.message = message

def get_config(config_file):
    with open(config_file, 'r') as toc_yaml:
        config = yaml.safe_load(toc_yaml)

    num_items = len(config['toc_items'])
    name = config.get('title', 'Copy of BASE')
    body = {
        'name': name
    }
    toc = "\n".join(config['toc_items'])
    if 'base_id' in config:
        base_id = config['base_id']
    else:
        raise ConfigurationError('Must specify a base presentation ID in the yaml file from which to load')
    return num_items, body, toc, base_id

def copy_base(slides, pres):
    base_page_id = slides.presentations().get(presentationId=pres).execute().get('slides')[0]["objectId"]
    body = {
                "requests": [{
                    "duplicateObject": {
                        "objectId": base_page_id,
                    }
                }]
            }
    response = slides.presentations().batchUpdate(presentationId=pres, body=body).execute()
    create_slide_response = response.get('replies')[0].get('duplicateObject')
    print('Created slide with ID: {0}'.format(create_slide_response.get('objectId')))

def update_base_slides (service, slides, toc, pres):
    reqs = []
    for i, slide in enumerate(slides):
        bodyId = slide['pageElements'][0]['objectId']
        first_half = "\n".join(toc.split("\n")[0:i + 1])
        second_half = "\n".join(toc.split("\n")[i+1:])
        reqs.extend([
                        {'insertText': {'objectId': bodyId, 'text': toc}},
                        {'updateTextStyle': {'objectId': bodyId,
                                                'style': {'bold': True},
                                                'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': len(first_half)},
                                                'fields': 'bold'},
                        },
                        {'updateTextStyle': {'objectId': bodyId,
                                                'style': {
                                                            'foregroundColor': {
                                                                'opaqueColor': {
                                                                    'rgbColor': {
                                                                        'red': 24/255,
                                                                        'blue': 76/255,
                                                                        'green': 53/255,
                                                                    }
                                                                }
                                                            }
                                                        },
                                                'textRange': {'type': 'FIXED_RANGE', 'startIndex': 0, 'endIndex': len(first_half)},
                                                'fields': 'foregroundColor'},
                        },
                        {'updateTextStyle': {'objectId': bodyId,
                                                'style': {
                                                            'foregroundColor': {
                                                                'opaqueColor': {
                                                                    'rgbColor': {
                                                                        'red': 137/255,
                                                                        'blue': 137/255,
                                                                        'green': 137/255,
                                                                    }
                                                                }
                                                            }
                                                        },
                                                'textRange': {'type': 'FROM_START_INDEX', 'startIndex': len(first_half)},
                                                'fields': 'foregroundColor'},
                        },
                        {
                            'createImage': {
                                'objectId': 'yellow_arrow' + str(i),
                                'url': IMG_URL,
                                'elementProperties': {
                                    'pageObjectId': slide['objectId'],
                                    'size': {
                                        'height': {'magnitude': 34725, 'unit': 'EMU'},
                                        'width': {'magnitude': 37650, 'unit': 'EMU'}
                                    },
                                    'transform': {
                                        'scaleX': 5.8811,
                                        'scaleY': 5.8812,
                                        'translateX': 348605.57,
                                        'translateY': 1006155.0125 + (i * 318000), # TODO: This is terrible, but I don't know how to get the translateYs of the bullet points
                                        'unit': 'EMU'
                                    },
                                },
                            }
                        },
                    ])
        
        reqs.extend([
                        {'createParagraphBullets': 
                            {'objectId': bodyId, 
                            'textRange': {'type': 'ALL'}, 
                            'bulletPreset': "NUMBERED_DIGIT_ALPHA_ROMAN"
                            }
                        },
                        ])
        
    service.presentations().batchUpdate(presentationId=pres, body={"requests": reqs}).execute()

def main(config_file="toc_example.yaml"):
    """
    Instantiates a slides_ and drive_service, copies over the base
    presentation into a new one with the specified name, 
    then generates a table of contents based on the provided .yaml file

    :param config_file: YAML file from which to load the Table of Contents
    """
    print("Getting credentials...")
    creds = None
    # The pickle file stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)   

    print(f"Initializing drive and slides service...")
    slides_service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    # Get configuration data
    num_toc_items, create_pres_request, toc_string, base_id = get_config(config_file)

    # Create new presentation 
    print(f"Copying base presentation...")
    drive_response = drive_service.files().copy(fileId=base_id, body=create_pres_request).execute()
    presentation_copy_id = drive_response.get('id')

    # Copy base slide for however many items you need
    print(f"Copying base slide x {num_toc_items} into new presentation with name {create_pres_request['name']}...")
    for i in range(num_toc_items - 1):
        copy_base(slides_service, presentation_copy_id)

    # Create actual table of contents content
    print("Populating slides with content from yaml file...")
    slides = slides_service.presentations().get(presentationId=presentation_copy_id).execute().get('slides')
    update_base_slides (slides_service, slides, toc_string, presentation_copy_id)
    print(f"New presentation with name {create_pres_request['name']} created.")

if __name__ == '__main__':
    clize.run(main)