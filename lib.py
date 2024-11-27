from dataclasses import dataclass
import uuid
import json
from dataclasses import dataclass
from datetime import datetime
from jira import JIRA
from getpass import getpass
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

def google_slides_connect():
    '''Authenticate to Google Slides and generate a service to be leveraged when updating/populating slides

    Returns: 
        (googleapiclient.discovery.Resource)
    '''
    
    flow = InstalledAppFlow.from_client_secrets_file(
        'google-credentials.json',
        scopes=['https://www.googleapis.com/auth/presentations'])
    
    credentials = flow.run_local_server()
    
    return build('slides', 'v1', credentials=credentials)

def updateSlides(service, **kwargs):
    '''General function for pushing updates to a specific slide deck'''
    return service.presentations().batchUpdate(**kwargs).execute()

@dataclass
class JiraRoadmapIssue:
    jira_id:str
    product_categories:list
    jira_quarter:str
    jira_link:str
    summary:str = ''
    description:str = '',
    beta:bool = False

def get_roadmap_issues(
        jira_service, 
        jira_project,
        issue_type, 
        product_category_mode, 
        product_category_prefix, 
        include_beta, 
        beta_attribute_name
    ):
    '''Retrieves all of the Roadmap Initiatives in Jira Cloud

    Args: 
        jira_service (JIRA): authenticated service used for interacting with jira
        jira_project (str): unique name of the jira project
        issue_type (str): name of the type of issues that represent roadmap items
        product_category_mode (str): 'components' or 'labels' 
        include_beta (bool): whether beta roadmap items are included
        beta_attribute_name (str): name of the attribute where the beta flag is held in jira
    '''
    
    jql_filter = f'project = {jira_project} and issuetype = "{issue_type}"'

    roadmap_issue_ids = jira_service.search_issues(
        jql_str=jql_filter,
        maxResults=None
    )

    if roadmap_issue_ids:
        
        roadmap_issues = []
        
        for issue_id in roadmap_issue_ids: 

            issue = jira_service.issue(id=issue_id)

            if product_category_mode == 'components':

                issue_categories = [comp.name for comp in issue.fields.components]
            
            elif product_category_mode == 'labels':
                
                issue_categories = issue.fields.labels

            else:
                raise Exception(f"product_category_mode: '{product_category_mode}' is not valid. Choose from either 'components' or 'labels'")

            filtered_categories = []

            for category in issue_categories:

                if category.startswith(product_category_prefix):

                    filtered_categories.append(category[len(product_category_prefix):])
            
            if not issue.fields.description: 

                issue.fields.description = ''

            beta_attr = getattr(issue.fields, beta_attribute_name)


            if include_beta and beta_attr and beta_attr.value == "Beta":

                beta_flag = True
            
            elif not include_beta and beta_attr and beta_attr.value == "Beta":
            
                continue

            else: 

                beta_flag = False

            roadmap_issues.append(JiraRoadmapIssue(
                summary=issue.fields.summary,
                description=issue.fields.description,
                jira_id=issue_id.id,
                product_categories=filtered_categories,
                jira_quarter=issue.fields.status.name,
                jira_link=issue.permalink(),
                beta = beta_flag
            ))
            
        return roadmap_issues

    else: 

        raise Exception("No jira issues were found with the provided JQL Filter")

def gen_header_slide_req(title):
    '''Creates a request body for a new slide

    Args: 
        title (Str): the title of the slide

    Returns: 
        (List(dict)): request body to generate roadmap item
        (str): Google Slides id of the object created
    '''

    titleId = str(uuid.uuid4())
    slideId = str(uuid.uuid4())

    request_body = [
        {
            'createSlide': {
                'objectId':slideId,
                'slideLayoutReference': {
                    'predefinedLayout': 'SECTION_HEADER'
                },
                'placeholderIdMappings': [
                    {
                        'objectId': titleId,
                        'layoutPlaceholder': {'type': 'TITLE', 'index': 0}
                    }
                ]
            }
        },
        {'insertText': {'objectId': titleId, 'text': title}},
    ]
    
    return request_body, slideId

def gen_roadmap_slide_req(
    title,
    roadmap_slide_config
):
    '''Creates a request body for a new slide

    Args: 
        title (Str): the title of the slide
        roadmap_slide_config (dict): configuration for the roadmap slides
        ...
        
    Returns: 
        (List(dict)): request body to generate roadmap item
        (str): Google Slides id of the object created
    '''

    left_header_config = roadmap_slide_config['left_header'] 
    left_main_config = roadmap_slide_config['left_main'] 
    timeline_arrow_config = roadmap_slide_config['timeline_arrow'] 
    quarter_marker_config = roadmap_slide_config['quarter_marker'] 
    quarters_config = roadmap_slide_config['quarters'] 
    roadmap_box_config = roadmap_slide_config['roadmap_box']
    columns_config = roadmap_slide_config['columns']
    
    title_id = str(uuid.uuid4())
    slide_id = str(uuid.uuid4())    
    left_header_element_id = str(uuid.uuid4())
    left_main_element_id = str(uuid.uuid4())

    left_main_locx=left_header_config["locx"]
    
    # timeline arrow
    timeline_arrow_start_locx = left_header_config["locx"] + left_header_config["width"]
    timeline_arrow_start_locy = left_header_config["locy"] + left_header_config["height"]/2
    timeline_arrow_element_id = str(uuid.uuid4())
    
    
    slide_req = [{
        'createSlide': {
            'objectId':slide_id,
            'slideLayoutReference': {
                'predefinedLayout': 'TITLE_AND_TWO_COLUMNS'
            },
            'placeholderIdMappings': [
                {
                    'objectId': title_id,
                    'layoutPlaceholder': {'type': 'TITLE', 'index': 0}
                }
            ]
        }
    }]

    title_req = [{'insertText': {'objectId': title_id, 'text': title}}]
    
    left_main_req = [{
            "createShape": {
                "objectId": left_main_element_id,
                "shapeType": left_main_config["shape_type"],
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "height": {"magnitude": left_main_config["height"], "unit": "PT"}, 
                        "width": {"magnitude": left_main_config["width"], "unit": "PT"}
                    },
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": left_main_locx,
                        "translateY": left_main_config["locy"],
                        "unit": "PT",
                    }
                }
            }
        },
        {
            "updateShapeProperties" : {
                "fields" : "contentAlignment, \
                    outline.outlineFill.solidFill.alpha, \
                    shapeBackgroundFill.solidFill.color.themeColor",
                "objectId": left_main_element_id,
                "shapeProperties" : {
                    "contentAlignment": "TOP",
                    "outline": {"outlineFill": {"solidFill":{"alpha":0}}},
                    "shapeBackgroundFill":{"solidFill": {"color": {"themeColor":left_main_config["color"]}}}
                }
            }
        },
        {
            "insertText": {
                "objectId": left_main_element_id,
                "insertionIndex": 0,
                "text": f"\n\n\n{left_main_config['text']}"
            }
        },
        {
            "updateTextStyle" : {
                "fields": "fontFamily, fontSize.magnitude, fontSize.unit",
                "objectId": left_main_element_id,
                "style" : {
                    "fontFamily": "Manrope",
                    "fontSize": {"magnitude": left_main_config["font_size"], "unit":"PT"}
                },
            }
        },
        {
            "updateParagraphStyle" : {
                "fields": "alignment",
                "objectId": left_main_element_id,
                "style" : {
                    "alignment": "START",
                },
            }
        }]
    
    left_header_req = [{
            "createShape": {
                "objectId": left_header_element_id,
                "shapeType": left_header_config["shape_type"],
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "height": {"magnitude": left_header_config["height"], "unit": "PT"}, 
                        "width": {"magnitude": left_header_config["width"], "unit": "PT"}
                    },
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": left_header_config["locx"],
                        "translateY": left_header_config["locy"],
                        "unit": "PT",
                    }
                }
            }
        },
        {
            "updateShapeProperties" : {
                "fields" : "contentAlignment, \
                    outline.outlineFill.solidFill.alpha, \
                    shapeBackgroundFill.solidFill.color.themeColor",
                "objectId": left_header_element_id,
                "shapeProperties" : {
                    "contentAlignment": "TOP",
                    "outline": {"outlineFill": {"solidFill":{"alpha":0}}},
                    "shapeBackgroundFill":{"solidFill": {"color": {"themeColor":left_header_config["color"]}}}
                }
            }
        },
        {
            "insertText": {
                "objectId": left_header_element_id,
                "insertionIndex": 0,
                "text": title
            }
        },
        {
            "updateTextStyle" : {
                "fields": "bold, fontFamily, fontSize.magnitude, fontSize.unit, \
                    foregroundColor.opaqueColor.themeColor",
                "objectId": left_header_element_id,
                "style" : {
                    "bold":True,
                    "fontFamily": "Manrope",
                    "fontSize": {"magnitude": left_header_config["font_size"], "unit":"PT"},
                    "foregroundColor": {"opaqueColor": {"themeColor":"LIGHT1"}}
                },
            }
        },
        {
            "updateParagraphStyle" : {
                "fields": "alignment",
                "objectId": left_header_element_id,
                "style" : {
                    "alignment": "CENTER",
                },
            }
        }]

    timeline_arrow_req = [{
            "createLine": {
                "objectId": timeline_arrow_element_id,
                "lineCategory": 'STRAIGHT',
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "height": {"magnitude": 1, "unit": "PT"}, 
                        "width": {"magnitude": timeline_arrow_config["width"], "unit": "PT"}
                    },
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": timeline_arrow_start_locx,
                        "translateY": timeline_arrow_start_locy,
                        "unit": "PT",
                    }
                }
            }
        },
        {
            "updateLineProperties" : {
                "fields" : "endArrow, lineFill.solidFill.color.themeColor, weight",
                "objectId": timeline_arrow_element_id,
                "lineProperties" : {
                    "endArrow": "FILL_ARROW",
                    "lineFill":{"solidFill": {"color": {"themeColor":timeline_arrow_config["color"]}}},
                    "weight": {"magnitude": timeline_arrow_config["weight"], "unit": "PT"}
                }
            }
        }]

    quarter_marker_reqs = []

    quarter_width = (timeline_arrow_config["width"] - roadmap_box_config['x_padding']*2) / len(columns_config)
    quarter_locx = timeline_arrow_start_locx + roadmap_box_config['x_padding']

    text_box_width = 200
    
    for col_num, column in enumerate(columns_config):
        
        quarter_marker_element_id = str(uuid.uuid4())
        quarter_textbox_element_id = str(uuid.uuid4())
    
        quarter_marker_reqs += [
            {
                "createShape": {
                    "objectId": quarter_marker_element_id,
                    "shapeType": quarter_marker_config["shape_type"],
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "height": {"magnitude": quarter_marker_config["height"], "unit": "PT"}, 
                            "width": {"magnitude": quarter_marker_config["width"], "unit": "PT"}
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": quarter_locx + quarter_width/2 - quarter_marker_config['width']/2 + (quarter_width * col_num),
                            "translateY": timeline_arrow_start_locy - quarter_marker_config["height"]/2,
                            "unit": "PT",
                        }
                    }
                }
            },
            {
                "updateShapeProperties" : {
                    "fields" : "outline.outlineFill.solidFill.alpha, \
                        shapeBackgroundFill.solidFill.color.themeColor",
                    "objectId": quarter_marker_element_id,
                    "shapeProperties" : {
                        "outline": {"outlineFill": {"solidFill":{"alpha":0}}},
                        "shapeBackgroundFill":{"solidFill": {"color": {"themeColor":quarter_marker_config["color"]}}}
                    }
                }
            },
            {
                "createShape": {
                    "objectId": quarter_textbox_element_id,
                    "shapeType": "RECTANGLE",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "height": {"magnitude": quarter_marker_config["height"], "unit": "PT"}, 
                            "width": {"magnitude": text_box_width, "unit": "PT"}
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": quarter_locx + quarter_width/2 - text_box_width/2 + (quarter_width * col_num),
                            "translateY": timeline_arrow_start_locy - quarter_marker_config["height"]*1.75,
                            "unit": "PT",
                        }
                    }
                }
            },
            {
                "updateShapeProperties" : {
                    "fields" : "outline.outlineFill.solidFill.alpha, \
                        shapeBackgroundFill.solidFill.alpha",
                    "objectId": quarter_textbox_element_id,
                    "shapeProperties" : {
                        "outline": {"outlineFill": {"solidFill":{"alpha":0}}},
                        "shapeBackgroundFill":{"solidFill":{"alpha":0}}
                    }
                }
            },
            {
                "insertText": {
                    "objectId": quarter_textbox_element_id,
                    "insertionIndex": 0,
                    "text": column['label']
                }
            },
            {
                "updateTextStyle" : {
                    "fields": "bold, fontFamily, fontSize.magnitude, foregroundColor.opaqueColor.themeColor",
                    "objectId": quarter_textbox_element_id,
                    "style" : {
                        "bold": True,
                        "fontFamily": quarter_marker_config['font'],
                        "fontSize": {"magnitude": quarter_marker_config["font_size"], "unit":"PT"},
                        "foregroundColor": {"opaqueColor": {"themeColor":quarter_marker_config['font_color']}}
                    }
                }
            }                
        ]
    
    request_body = slide_req + title_req + left_main_req + left_header_req + timeline_arrow_req + quarter_marker_reqs
    
    return request_body, slide_id

def gen_roadmap_item_req(page_id, width, locx, locy, roadmap_box_config, tagline, description, link, beta):
    '''Creates the necessary request body for generating a roadmap item in Google Slides

    Args: 
        page_id (str): id of the Google Slide page
        width (int): pixel width of roadmap box
        locx (int): location on x-axis
        loxy (int): location on y-axis
        tagline (str): short tagline for the roadmap item
        roadmap_box_config (dict): configuration of the roadmap boxes
        description (str): 1 - 2 sentences describing the roadmap intiative
        link (str): URL to the jira roadmap item
        beta (bool): is the roadmap item a beta initiative

    Returns: 
        (List(dict)): request body to generate roadmap item
        (str): Google Slides id of the object created
    '''
    
    # Properties for all roadmap item boxes
    SHAPE_TYPE='ROUND_RECTANGLE'
    
    ptHeight = {"magnitude": roadmap_box_config["height"], "unit": "PT"}
    ptWidth = {"magnitude": width, "unit": "PT"}

    element_id = str(uuid.uuid4())
    
    request_body = [
        {
            "createShape": {
                "objectId": element_id,
                "shapeType": roadmap_box_config["shape_type"],
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": {"height": ptHeight, "width": ptWidth},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": locx,
                        "translateY": locy,
                        "unit": "PT",
                    }
                }
            }
        },
        {
            "updateShapeProperties" : {
                "fields" : "contentAlignment, \
                    outline.outlineFill.solidFill.color.themeColor, \
                    link.url, \
                    shapeBackgroundFill.solidFill.color.themeColor",
                "objectId": element_id,
                "shapeProperties" : {
                    "contentAlignment": "TOP",
                    "outline": {"outlineFill": {"solidFill": {"color": {"themeColor":roadmap_box_config["outline_color"]}}}},
                    "link": {"url": link},
                    "shapeBackgroundFill":{"solidFill": {"color": {"themeColor":roadmap_box_config["fill_color"]}}}
                }
            }
        },
        {
            "insertText": {
                "objectId": element_id,
                "insertionIndex": 0,
                "text": tagline + '\n' + description,
            }
        },
        {
            "updateTextStyle" : {
                "fields": "bold, fontFamily, fontSize.magnitude, fontSize.unit",
                "objectId": element_id,
                "style" : {
                    "fontFamily": "Manrope",
                    "fontSize": {"magnitude": roadmap_box_config["font_size"], "unit":"PT"}
                },
            }
        },
        {
            "updateParagraphStyle" : {
                "fields": "alignment",
                "objectId": element_id,
                "style" : {
                    "alignment": "START",
                    #"indentStart": {"magnitude": -0.1, "unit":"PT"},
                    #"indentEnd": {"magnitude": width-0.1, "unit":"PT"}
                },
            }
        },
        {
            "updateTextStyle" : {
                "fields": "bold",
                "objectId": element_id,
                "style" : {
                    "bold": True,
                },
                "textRange": {"endIndex": len(tagline), "startIndex": 0, "type": "FIXED_RANGE"}
            }
        }
    ]

    if beta: 

        beta_flag_element_id = str(uuid.uuid4())

        beta_flag_req_body = [
            {
                "createShape": {
                    "objectId": beta_flag_element_id,
                    "shapeType": "FLOW_CHART_TERMINATOR",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {
                            "height": {"magnitude": roadmap_box_config["height"]/4.5, "unit": "PT"}, 
                            "width": {"magnitude": width/7, "unit": "PT"}
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": locx + width * 6/7 - 2,
                            "translateY": locy + 2,
                            "unit": "PT",
                        }
                    }
                }
            },
            {
                "updateShapeProperties" : {
                    "fields" : "contentAlignment, \
                        outline.outlineFill.solidFill.color.themeColor, \
                        shapeBackgroundFill.solidFill.color.themeColor",
                    "objectId": beta_flag_element_id,
                    "shapeProperties" : {
                        "contentAlignment": "MIDDLE",
                        "outline": {"outlineFill": {"solidFill": {"color": {"themeColor":roadmap_box_config["beta_label_outline_color"]}}}},
                        "shapeBackgroundFill":{"solidFill": {"color": {"themeColor":roadmap_box_config["beta_label_color"]}}}
                    }
                }
            },
            {
                "insertText": {
                    "objectId": beta_flag_element_id,
                    "insertionIndex": 0,
                    "text": "beta",
                }
            },
            {
                "updateTextStyle" : {
                    "fields": "bold, fontFamily, fontSize.magnitude, fontSize.unit, foregroundColor.opaqueColor.themeColor",
                    "objectId": beta_flag_element_id,
                    "style" : {
                        "fontFamily": "Manrope",
                        "fontSize": {"magnitude": roadmap_box_config["beta_label_font_size"], "unit":"PT"},
                        "foregroundColor": {"opaqueColor": {"themeColor":roadmap_box_config['beta_label_font_color']}}
                    },
                }
            },
            {
                "updateParagraphStyle" : {
                    "fields": "alignment",
                    "objectId": beta_flag_element_id,
                    "style" : {
                        "alignment": "CENTER",
                    },
                }
            }
        ]

        request_body += beta_flag_req_body

    return request_body, element_id

def get_unique_product_groups(roadmap_issues):

    categories = []
    
    for issue in roadmap_issues:

        categories += issue.product_categories

    return list(set(categories))

@dataclass
class RoadmapSlide:
    title:str
    google_slide_id:str
    product_category:str

def generate_roadmap_slides(presentation_id, slides_service, product_categories, roadmap_slide_config):
    '''Creates the section headers and placeholder roadmap slides to populate with roadmap items

    Args:
        presentation_id (str): google id of presentation to add slides to
        slides_service (): google authenticated service to use when creating slides
        product_categories (list(str)): list of unique product categories
        roadmap_slide_config (dict): all of the properties required to create the roadmap slides

    Returns:
        (list(RoadmapSlides)): slides generated 
    '''
    
    slide_reqs = []
    slides = []
    
    for category in product_categories:
    
        header_slide_req, header_slide_id = gen_header_slide_req(title=category)
    
        slide_reqs += header_slide_req    
    
        roadmap_slide_req, roadmap_slide_id = gen_roadmap_slide_req(
            title = category, 
            roadmap_slide_config = roadmap_slide_config
        )
    
        slides.append(RoadmapSlide(
            title=category,
            google_slide_id=roadmap_slide_id,
            product_category=category
        ))
        
        slide_reqs += roadmap_slide_req

    res = updateSlides(
        service = slides_service,
        presentationId = presentation_id,
        body = {'requests': slide_reqs}
    )

    return slides

def populate_roadmap_with_issues(
    presentation_id, 
    slides_service, 
    roadmap_slides, 
    roadmap_slide_config,
    jira_roadmap_issues
):
    '''Places the roadmap items on the slides provided based on the config

    Args:
        presentation_id (str): google id of presentation to add slides to
        slides_service (): google authenticated service to use when creating slides
        roadmap_slides (list(Slide)): slides for placing the roadmap items on
        roadmap_slide_config (dict): configuraiton for the roadmap slides
        jira_roadmap_issues (list(JiraRoadmapIssues)): list of roadmap issues to add to the slides

    Returns: 
        (list): updates made to the slides
    '''

    roadmap_box_config = roadmap_slide_config['roadmap_box']
    arrow_config = roadmap_slide_config['timeline_arrow']
    left_header_config = roadmap_slide_config['left_header']
    columns_config = roadmap_slide_config['columns']

    roadmap_box_width = (arrow_config['width'] - roadmap_box_config['x_padding'] * (len(columns_config) + 1)) \
        / len(columns_config)
    roadmap_box_locx = left_header_config['locx'] + left_header_config['width'] + roadmap_box_config['x_padding']
    roadmap_box_locy = left_header_config['locy'] + left_header_config['height']
    
    roadmap_shapes = []

    for slide in roadmap_slides:

        for col in columns_config:
            col['count'] = 0
    
        for issue in jira_roadmap_issues:
    
            if slide.product_category in issue.product_categories:

                for col_num, col in enumerate(columns_config):
                    
                    if issue.jira_quarter in col['jira_statuses']:

                        locx = roadmap_box_locx + (roadmap_box_width + roadmap_box_config["x_padding"]) * col_num

                        locy = col['count'] * (roadmap_box_config["height"] + roadmap_box_config["y_padding"]) + roadmap_box_locy

                        col['count'] += 1

                        roadmap_shape, shape_id = gen_roadmap_item_req(
                            page_id=slide.google_slide_id,
                            tagline=issue.summary,
                            description=issue.description[:roadmap_box_config["description_length"]],
                            width=roadmap_box_width,
                            locx=locx,
                            locy=locy,
                            link=issue.jira_link,
                            roadmap_box_config=roadmap_box_config,
                            beta = issue.beta
                        )
                
                        roadmap_shapes.append(roadmap_shape)
    
    return updateSlides(
        service = slides_service,
        presentationId = presentation_id,
        body = {'requests': roadmap_shapes}
    )

def generate_roadmap_deck(jira_service, google_service, roadmap_slide_config, presentation_id):
    '''Get the jira roadmap issues and generate all of the slides with the details
    
    Args: 
        jira_service (JIRA): authenticated service used for interacting with jira
        google_service (): google authenticated service to use when creating slides
        roadmap_slide_config (dict): configuraiton for the roadmap slides 
        presentation_id (str): google id of presentation to add slides to
    '''

    jira_roadmap_issues = get_roadmap_issues(
        jira_service=jira_service, 
        **roadmap_slide_config['jira_roadmap_issues']
    )
    
    product_categories = get_unique_product_groups(jira_roadmap_issues)
    
    slides = generate_roadmap_slides(
        slides_service=google_service, 
        presentation_id=presentation_id, 
        product_categories=product_categories,
        roadmap_slide_config=roadmap_slide_config
    )
    
    res_populate_roadmap = populate_roadmap_with_issues(
        presentation_id=presentation_id,
        slides_service=google_service,
        roadmap_slides=slides, 
        roadmap_slide_config=roadmap_slide_config,
        jira_roadmap_issues=jira_roadmap_issues
    )

    return f"Generated {len(slides)*2} slides, {len(product_categories)} product categories, and {len(jira_roadmap_issues)} roadmap items."