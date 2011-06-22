from suds import WebFault
from suds.client import Client
import os

class LiveDocx():

    API_URL = 'https://api.livedocx.com/1.2/mailmerge.asmx?wsdl'
    ALLOWED_TEMPLATE_EXT = ['DOC', 'DOCX', 'RTF', 'TXD']
    ALLOWED_DOCUMENT_EXT = ['DOC', 'DOCX', 'HTML', 'PDF', 'TXD', 'TXT']
    ALLOWED_IMAGE_EXT = ['BMP', 'GIF', 'JPG', 'PNG', 'TIFF']

    def __init__(self):
        self.field_values = {}
        self.block_field_values = {}
        self.client = Client(self.API_URL)

    def assign_value(self, key, value):
        """
        Assing value to single template variable
        """
        self.field_values[key] = value

    def assign_block(self, key, values):
        """
        Assign block value for template block. Ex.:
        l = LiveDocx()
        l.assign_block('names', [
            {'first': 'John'},
            {'first': 'Sam'}
        ]
        """
        self.block_field_values[key] = values

    def create_document(self):
        """
        Creates document with assigned variables, clear them and preparing document to retrieve.
        Call retrieve_document() next.
        """
        # set single values

        if len(self.field_values) > 0:
            self._set_field_values()

        # set multi values
        if len(self.block_field_values) > 0:
            self._set_multi_field_values()

        self.field_values = {}
        self.block_field_values = {}

        self.client.service.CreateDocument()


    def delete_template(self, filename):
        """
        Delete template from server.
        If template with that name doesn't exists, throws LiveDocxError
        """
        if self.template_exists(filename):
            self.client.service.DeleteTemplate(filename=filename)
        else:
            raise LiveDocxError('Template "%s" not exists and it cannot be deleted' % filename)

    def download_template(self, filename):
        """
        Download selected template from server.
        Returns binary data from template
        """
        if self.template_exists(filename):
            return self.client.service.DownloadTemplate(filename).decode('base64')
        else:
            raise LiveDocxError('Template "%s" not exists' % filename)

    def get_bitmaps(self, zoom, format, from_page=None, to_page=None):
        """
        Returns a list containing bitmaps binary data for the specified pages.

        Arguments:
        zoom        -- zoom factor from 20 to 400
        format      -- format of recieved images
        from_page   -- start page to extract images
        to_page     -- end page to extract images

        If from_page/to_page is set, second value also has to be set. In another case all bitmaps are returned
        """
        if zoom < 20 or zoom > 400:
            raise LiveDocxError('Zoom factor value must be between 20 and 400')

        if format not in self.ALLOWED_IMAGE_EXT:
            raise LiveDocxError('Invalid image format. Valid formats are: ' + repr(self.ALLOWED_IMAGE_EXT))

        if from_page is not None and to_page is not None:
            bitmaps = self.client.service.GetBitmaps(fromPage = from_page, toPage = to_page, zoomFactor=zoom, format=format)
        elif from_page is None and to_page is None:
            bitmaps = self.client.service.GetAllBitmaps(zoomFactor=zoom, format=format)
        else:
            raise LiveDocxError('Both values from_page and to_page must be set')

        return bitmaps.string

    def get_metafiles(self, from_page=None, to_page=None):
        """
        Returns a metafile (i.e. the image format) string array from the specified pages.

        Arguments:
        from_page   -- start page to extract images
        to_page     -- end page to extract images

        If from_page/to_page is set, second value also has to be set. In another case all metadatas are returned
        """
        if from_page is not None and to_page is not None:
            data = self.client.service.GetMetafiles(fromPage=from_page, toPage=to_page)
        elif from_page is None and to_page is None:
            data = self.client.service.GetAllMetafiles()
        else:
            raise LiveDocxError('Both values from_page and to_page must be set')

        return self._parse_response(data)

    def get_block_names(self):
        """
        Returns the merge block names.
        """
        return self._parse_response(self.client.service.GetBlockNames())

    def get_field_names(self):
        """
        Returns the merge block names.
        """
        return self._parse_response(self.client.service.GetFieldNames())

    def get_font_names(self):
        """
        Returns a list of all fonts on the respective LiveDocx server that can be used with the API.
        """
        return self._parse_response(self.client.service.GetFontNames())

    def list_templates(self):
        """
        Returns list of templates stored on server.
        Each elements is described as dictionary
        """
        templates_data = self.client.service.ListTemplates()
        return [
        {
            'name': template.string[0],
            'size': template.string[1],
            'created_at': template.string[2],
            'modified_at': template.string[3]
        }
        for template in templates_data.ArrayOfString]


    def login(self, username, password):
        """
        Tries to login into service and throws LiveDocxError on invalid creditnetials
        """
        try:
            self.client.service.LogIn(username=username, password=password)
        except WebFault:
            raise LiveDocxError('Invalid username/password combination')

    def logout(self):
        """
        Logout from serivce
        """
        self.client.service.LogOut()

    def retrieve_document(self, format):
        """
        Returns document binary data in specyfict format

        Arguments:
        format  -- one of the allowed formats: ['DOC', 'DOCX', 'HTML', 'PDF', 'TXD', 'TXT']
        """
        self._validate_extension(format.upper(), self.ALLOWED_DOCUMENT_EXT)
        return self.client.service.RetrieveDocument(format=format.upper()).decode('base64')

    def set_ignore_sub_templates(self):
        """
        Specifies whether INCLUDETEXT field should be merged.
        """
        self.client.service.SetIgnoreSubTemplates()

    def set_local_template(self, filename):
        """
        Specifies a local template that can be used for further processing.
        This template WILL NOT be stored on server after closing connection.

        Argumens:
        filename    -- path to template filename (must by in valid format: ['DOC', 'DOCX', 'RTF', 'TXD'])

        """
        extension = self._get_ext(filename)
        self._validate_extension(extension.upper(), self.ALLOWED_TEMPLATE_EXT)

        template = open(filename).read().encode('base64')

        self.client.service.SetLocalTemplate(template=template, format=extension.upper())

    def template_exists(self, filename):
        """
        Checks if template stored on server

        Arguments:
        filename    -- remote template name
        """
        return self.client.service.TemplateExists(filename=filename)

    def set_remote_template(self, filename):
        """
        Chooses remote template for creating document

        Arguments:
        filename    -- remote template name

        Throws an LiveDocxError when template don't exists
        """
        if self.template_exists(filename):
            self.client.service.SetRemoteTemplate(filename=filename)
        else:
            raise LiveDocxError('Remote template "%s" not exists' % filename)

    def upload_template(self, path, filename):
        """
        Uploads template to the server

        Arguments:

        path        -- local path with template (must be in valid format: ['DOC', 'DOCX', 'RTF', 'TXD'])
        filename    -- template name on server (must be in valid format: ['DOC', 'DOCX', 'RTF', 'TXD'])

        Throws an exception when
        """
        template_data = open(path).read().encode('base64')
        local_ext = self._get_ext(path)

        self._validate_extension(local_ext.upper(), self.ALLOWED_TEMPLATE_EXT)

        if local_ext != self._get_ext(filename):
            raise LiveDocxError('Local template extension and remote name extension must match')

        self.client.service.UploadTemplate(template=template_data, filename=filename)

    def __dell__(self):
        self.logout()

    def _create_soap_object(self, name):
        """
        Creates complex object using built in factory
        """
        return self.client.factory.create(name)

    def _parse_response(self, response):
        """
        Extract data from response if needed
        """
        if response is not None:
            return response.string
        return response

    def _set_field_values(self):
        """
        Sets values for single fields
        """
        data = self._create_soap_object('ArrayOfArrayOfString')

        arr1 = self._create_soap_object('ArrayOfString')
        arr1.string = self.field_values.keys()

        arr2 = self._create_soap_object('ArrayOfString')
        arr2.string = self.field_values.values()

        data.ArrayOfString.append(arr1)
        data.ArrayOfString.append(arr2)

        self.client.service.SetFieldValues(fieldValues=data)

    def _set_multi_field_values(self):
        """
        Sets values for blocks
        """
        for block_key, block in self.block_field_values.items():

            data = self._create_soap_object('ArrayOfArrayOfString')

            names = self._create_soap_object('ArrayOfString')
            names.string = block[0].keys()
            data.ArrayOfString.append(names)

            for item in block:
                row = self._create_soap_object('ArrayOfString')
                row.string = item.values()
                data.ArrayOfString.append(row)

            self.client.service.SetBlockFieldValues(blockName=block_key, blockFieldValues=data)

    def _get_ext(self, path):
        """
        Extract extension from filename
        """
        return os.path.splitext(path)[1][1:]

    def _validate_extension(self, extension, allowed_extensions):
        """
        Check if extension has valid type based on type (template/document)
        """
        if extension not in allowed_extensions:
            raise LiveDocxError('That format isn\'t allowed - please pick one of these: '
            + repr(self.ALLOWED_DOCUMENT_EXT))

class LiveDocxError(Exception):
    """Base class for LiveDocx errors in the module"""
    pass