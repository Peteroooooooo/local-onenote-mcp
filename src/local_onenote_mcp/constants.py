"""OneNote COM constants used by the local MCP server."""

ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
XML_SCHEMA_2013 = 2

HIERARCHY_SCOPES = {
    "self": 0,
    "children": 1,
    "notebooks": 2,
    "sections": 3,
    "pages": 4,
}

CREATE_FILE_TYPES = {
    "none": 0,
    "notebook": 1,
    "folder": 2,
    "section_group": 2,
    "section": 3,
}

NEW_PAGE_STYLES = {
    "default": 0,
    "blank_with_title": 1,
    "blank_no_title": 2,
}

PAGE_INFO = {
    "basic": 0,
    "binary": 1,
    "selection": 2,
    "binary_selection": 3,
    "file_type": 4,
    "binary_file_type": 5,
    "selection_file_type": 6,
    "all": 7,
}

PUBLISH_FORMATS = {
    "one": 0,
    "onepkg": 1,
    "mhtml": 2,
    "mht": 2,
    "pdf": 3,
    "xps": 4,
    "word": 5,
    "doc": 5,
    "docx": 5,
    "emf": 6,
    "html": 7,
    "one2007": 8,
}

SPECIAL_LOCATIONS = {
    "backup": 0,
    "unfiled": 1,
    "default_notebook_folder": 2,
}

FILING_LOCATIONS = {
    "email": 0,
    "contacts": 1,
    "tasks": 2,
    "meetings": 3,
    "web_content": 4,
    "printouts": 5,
}

FILING_LOCATION_TYPES = {
    "named_section_new_page": 0,
    "current_section_new_page": 1,
    "current_page": 2,
    "named_page": 4,
}
