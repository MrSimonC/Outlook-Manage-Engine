import os
import requests
import xml.etree.ElementTree as ET
__version__ = 0.1


class API:
    """
    The main sending class for the Manage Service Engine Rest API
    # https://www.manageengine.com/products/service-desk/help/adminguide/api/request-operations.html
    """
    def __init__(self, api_key, api_url_base):
        self.api_key = api_key
        self.api_url_base = api_url_base

    @staticmethod
    def create_xml(fields):
        xml_string = ET.Element('Operation')
        details = ET.SubElement(xml_string, 'Details')
        for key, value in fields.items():
            param = ET.SubElement(details, 'parameter')
            ET.SubElement(param, 'name').text = key
            ET.SubElement(param, 'value').text = value
        return ET.tostring(xml_string)

    def send(self, url_append, operation, input_data_xml='', attachment=''):
        """
        Send through details into API
        :param url_append: string to append to end of base API url e.g. 21 but not /21
        :param operation: operation name param as specified in ManageEngine API spec
        :param input_data_xml: xml string to pass into API
        :param attachment: file path to attachment
        :return: {'response_key': 'response value', ...}
        """
        params = {'TECHNICIAN_KEY': self.api_key,
                  'OPERATION_NAME': operation}
        if input_data_xml:
            params.update({'INPUT_DATA': input_data_xml})
        if attachment:
            file = {'file': open(attachment, 'rb')}
            response_text = requests.post(os.path.join(self.api_url_base, url_append), params=params, files=file).text
        else:
            response_text = requests.get(os.path.join(self.api_url_base, url_append), params).text
        # print(response_text)
        response = ET.fromstring(response_text)
        result = {}
        for status_item in response.iter('result'):
            result = {
                'response_status': status_item.find('status').text,
                'response_message': status_item.find('message').text
            }
        if result['response_status'] == 'Success':
            for param_tags in response.iter("Details"):
                result.update(dict([(details_params.find('name').text, details_params.find('value').text)
                                    for details_params in param_tags.findall('parameter')
                                    if details_params is not None]))
        return result


def eg_view_request():
    api = API('16EE6838-8160-4EFC-AEC1-0B35A59AF42C', 'http://sdplus/sdpapi/request/')
    result = api.send('154594', 'GET_REQUEST')
    print(result)


def eg_add_request():
    api = API('16EE6838-8160-4EFC-AEC1-0B35A59AF42C', 'http://sdplus/sdpapi/request/')
    xml = api.create_xml({
        'reqtemplate': 'Default Request',  # or 'General IT Request' which has supplier ref, but also Due by Date,
        'requesttype': 'Service Request',
        'status': 'Hold - Awaiting Third Party',
        'requester': 'Simon Crouch',
        'mode': '@Southmead Retained Estate',  # site
        'best contact number': '-',
        'Exact Location': 'white room',
        'group': 'Back Office',
        'subject': 'This is a test call only - please ignore',
        'description': 'This is a test call only (description) - please ignore',
        'service': '.Lorenzo/Galaxy - IT Templates',    # Service Category
        'category': 'Clinical Applications Incident',  # Self Service Incident
        'subcategory': 'Lorenzo',
        'impact': '3 Impacts Department',
        'urgency': '3 Business Operations Slightly Affected - Requires response within 8 hours of created time'
    })
    result = api.send('', 'ADD_REQUEST', xml)
    print(result)


def eg_add_attachment():
    api = API('16EE6838-8160-4EFC-AEC1-0B35A59AF42C', 'http://sdplus/sdpapi/request/')
    result = api.send('137216/attachment', 'ADD_ATTACHMENT', attachment=r'C:\Users\nbf1707\Desktop\test.txt')
    print(result)


def eg_edit_request():
    api = API('16EE6838-8160-4EFC-AEC1-0B35A59AF42C', 'http://sdplus/sdpapi/request/')
    xml = api.create_xml({'subject': 'EDITED EDITED ignore this request'})
    result = api.send('154820', 'EDIT_REQUEST', xml)
    print(result)


def eg_delete_request():
    api = API('16EE6838-8160-4EFC-AEC1-0B35A59AF42C', 'http://sdplus/sdpapi/request/')
    result = api.send('154820', 'DELETE_REQUEST')
    print(result)
