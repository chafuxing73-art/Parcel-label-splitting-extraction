import pytest
import os
import tempfile
from app import app


class TestAppRoutes:
    def setup_method(self):
        app.config['TESTING'] = True
        app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
        app.config['OUTPUT_FOLDER'] = tempfile.mkdtemp()
        self.client = app.test_client()

    def test_index_route(self):
        response = self.client.get('/')
        assert response.status_code == 200

    def test_process_route_no_file(self):
        response = self.client.post('/api/process')
        assert response.status_code == 400
        data = response.get_json()
        assert data['success'] is False
        assert '没有文件' in data['error']

    def test_process_route_empty_filename(self):
        data = {'file': (b'', '')}
        response = self.client.post('/api/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400

    def test_process_route_invalid_file(self):
        data = {'file': ('test.txt', b'not a pdf', 'text/plain')}
        response = self.client.post('/api/process', data=data, content_type='multipart/form-data')
        assert response.status_code == 400
        json_data = response.get_json()
        assert json_data['success'] is False
        assert '请选择PDF文件' in json_data['error']
