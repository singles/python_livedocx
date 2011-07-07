from suds import WebFault
from suds.client import Client
import os

class LiveDocx(object):

    API_URL = 'https://api.livedocx.com/1.2/mailmerge.asmx?wsdl'
    ALLOWED_TEMPLATE_EXT = ['DOC', 'DOCX', 'RTF', 'TXD']
    ALLOWED_DOCUMENT_EXT = ['DOC', 'DOCX', 'HTML', 'PDF', 'TXD', 'TXT']
    ALLOWED_IMAGE_EXT = ['BMP', 'GIF', 'JPG', 'PNG', 'TIFF']

    def __init__(self):
        self.field_values = {}
        self.block_field_values = {}
        self.client = Client(self.API_URL)

    def assign_value(self, key, value):
        """Assign value to single template variable """
        self.field_values[key] = value

    def assign_block(self, key, values):
        """Assign block value for template block. Ex.:
        >>> l = LiveDocx()
        >>> l.assign_block('names', [
        >>>     {'first': 'John'},
        >>>     {'first': 'Sam'}
        >>> ]
        """
        self.block_field_values[key] = values

    def __setitem__(self, key, value):
        """Shortcut for assign_value/assign_values methods.
        >>> doc['foo'] = bar
        """
        if isinstance(key, (list, tuple)):
            self.assign_block(key, value)
        else:
            self.assign_value(key, value)

    def assign(self, data):
        """Assign data, and automatically add as and value or block.
        :param data: dictionary containing data
        >>> ld_object.assign({
        >>>     'name': 'John'
        >>>     'age': 123
        >>>     'data': [dict(field='foo'), dict(field='bar)]
        >>> })
        """
        if type(data) != dict:
            raise ValueError('Passed data must be a dictionary')

        for key, value in data.iteritems():
            if type(value) == list:
                self.assign_block(key, value)
            else:
                self.assign_value(key, value)

    def create_document(self):
        """Create document with assigned variables, clear them and preparing document to retrieve.
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
        """Delete template from server.
        :raise LiveDocxError: when template with that name doesn't exists
        """
        if self.template_exists(filename):
            self.client.service.DeleteTemplate(filename=filename)
        else:
            raise LiveDocxError('Template "%s" not exists and it cannot be deleted' % filename)

    def download_template(self, filename):
        """Download selected template from server.
        Returns binary data from template
        """
        if self.template_exists(filename):
            return self.client.service.DownloadTemplate(filename).decode('base64')
        else:
            raise LiveDocxError('Template "%s" not exists' % filename)

    def get_bitmaps(self, zoom, format, pages=None):
        """Return a list containing bitmaps binary data for the specified pages.
        :param zoom: zoom factor from 20 to 400
        :param format: format of recieved images
        :param pages: range of pages
        :type pages: tuple
        """
        if not (20 <= zoom <= 400):
            raise LiveDocxError('Zoom factor value must be between 20 and 400')
        if format not in self.ALLOWED_IMAGE_EXT:
            raise LiveDocxError('Invalid image format. Valid formats are: %s' % (', '.join(self.ALLOWED_IMAGE_EXT,)))

        if pages == None:
            bitmaps = self.client.service.GetAllBitmaps(zoomFactor=zoom, format=format)
        elif not None in pages:
            bitmaps = self.client.service.GetBitmaps(fromPage = pages[0], toPage = pages[1], zoomFactor=zoom, format=format)
        else:
            raise LiveDocxError('Both values from_page and to_page must be set')

        return bitmaps.string

    def get_metafiles(self, pages=None):
        """Return a metafile (i.e. the image format) from the specified pages.
        :param pages: range of pages to extract images
        :type pages: tuple
        :rtype list of strings
        """
        if pages == None:
            data = self.client.service.GetAllMetafiles()
        elif not None in pages:
            data = self.client.service.GetMetafiles(fromPage = pages[0], toPage = pages[1])
        else:
            raise LiveDocxError('Both values from_page and to_page must be set')

        return self._parse_response(data)

    def get_block_names(self):
        """Return the merge block names."""
        return self._parse_response(self.client.service.GetBlockNames())

    def get_field_names(self):
        """Return the merge block names."""
        return self._parse_response(self.client.service.GetFieldNames())

    def get_font_names(self):
        """Return a list of all fonts on the respective LiveDocx server that can be used with the API.
        """
        return self._parse_response(self.client.service.GetFontNames())

    def list_templates(self):
        """Return list of templates stored on server.
        :rtype list of dicts
        """
        templates_data = self.client.service.ListTemplates()
        return [
                {
                    'name': template.string[0],
                    'size': template.string[1],
                    'created_at': template.string[2],
                    'modified_at': template.string[3]
                } for template in templates_data.ArrayOfString
        ]


    def login(self, username, password):
        """Try to login into service
        :raise LiveDocxError: when invalid creditnetials given
        """
        try:
            self.client.service.LogIn(username=username, password=password)
        except WebFault:
            raise LiveDocxError('Invalid username/password combination')

    def logout(self):
        """Logout from serivce"""
        self.client.service.LogOut()

    def retrieve_document(self, format):
        """Return document binary data in specyfict format
        :param format: one of the allowed formats: ['DOC', 'DOCX', 'HTML', 'PDF', 'TXD', 'TXT']
        """
        self._validate_extension(format.upper(), self.ALLOWED_DOCUMENT_EXT)
        return self.client.service.RetrieveDocument(format=format.upper()).decode('base64')

    def set_ignore_sub_templates(self):
        """Specify whether INCLUDETEXT field should be merged."""
        self.client.service.SetIgnoreSubTemplates()

    def set_local_template(self, filename):
        """Specify a local template that can be used for further processing.
        This template WILL NOT be stored on server after closing connection.
        :param filename: path to template filename (must by in valid format: ['DOC', 'DOCX', 'RTF', 'TXD'])
        """
        extension = self._get_ext(filename)
        self._validate_extension(extension.upper(), self.ALLOWED_TEMPLATE_EXT)

        template = open(filename).read().encode('base64')

        self.client.service.SetLocalTemplate(template=template, format=extension.upper())

    def template_exists(self, filename):
        """Check if template stored on server
        :param filename: remote template name
        """
        return self.client.service.TemplateExists(filename=filename)

    def set_remote_template(self, filename):
        """Choose remote template for creating document
        :param filename: remote template name
        :raise LiveDocxError: when template doesn't exist
        """
        if self.template_exists(filename):
            self.client.service.SetRemoteTemplate(filename=filename)
        else:
            raise LiveDocxError('Remote template "%s" not exists' % filename)

    def upload_template(self, path, filename):
        """Upload template to the server
        :param path: local path with template (must be in valid format: ['DOC', 'DOCX', 'RTF', 'TXD'])
        :param filename: template name on server (must be in valid format: ['DOC', 'DOCX', 'RTF', 'TXD'])
        :raise LiveDocxError: when files' extensions differ
        """
        template_data = open(path).read().encode('base64')
        local_ext = self._get_ext(path)

        self._validate_extension(local_ext.upper(), self.ALLOWED_TEMPLATE_EXT)

        if local_ext != self._get_ext(filename):
            raise LiveDocxError('Local template extension and remote name extension must match')

        self.client.service.UploadTemplate(template=template_data, filename=filename)

    def __del__(self):
        self.logout()

    def _create_soap_object(self, name):
        """Create complex object using built in factory"""
        return self.client.factory.create(name)

    def _parse_response(self, response):
        """Extract data from response if needed"""
        if response is not None:
            return response.string
        return response

    def _set_field_values(self):
        """Set values for single fields"""
        data = self._create_soap_object('ArrayOfArrayOfString')

        arr1 = self._create_soap_object('ArrayOfString')
        arr1.string = self.field_values.keys()

        arr2 = self._create_soap_object('ArrayOfString')
        arr2.string = self.field_values.values()

        data.ArrayOfString.append(arr1)
        data.ArrayOfString.append(arr2)

        self.client.service.SetFieldValues(fieldValues=data)

    def _set_multi_field_values(self):
        """Set values for blocks"""
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
        """Extract extension from filename"""
        return os.path.splitext(path)[1][1:]

    def _validate_extension(self, extension, allowed_extensions):
        """Check if extension has valid type based on type (template/document)"""
        if extension not in allowed_extensions:
            raise LiveDocxError("That format isn't allowed - please pick one of these: %s" % (','.join(self.ALLOWED_TEMPLATE_EXT))

class LiveDocxError(Exception):
    """Base class for LiveDocx errors in the module"""
    pass
