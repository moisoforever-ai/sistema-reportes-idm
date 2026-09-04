import unittest
import json
from server import app

class TestAuthAndReactive(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True

    def test_unauthorized_access(self):
        # 1. Main page should redirect to login page
        response = self.app.get('/')
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith('/login'))

        # 2. API options should return 401
        response = self.app.get('/api/options')
        self.assertEqual(response.status_code, 401)
        data = response.get_json()
        self.assertFalse(data['success'])
        self.assertIn('No autorizado', data['error'])

        # 3. API generate should return 401
        response = self.app.post('/api/generate', json={})
        self.assertEqual(response.status_code, 401)

    def test_failed_login(self):
        response = self.app.post('/login', json={"username": "admin", "password": "wrongpassword"})
        self.assertEqual(response.status_code, 401)
        data = response.get_json()
        self.assertFalse(data['success'])
        self.assertEqual(data['error'], "Usuario o contraseña incorrectos")

    def test_successful_login_and_logout(self):
        # 1. Log in
        response = self.app.post('/login', json={"username": "idm", "password": "idm2026"})
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])

        # 2. Access index now (should be 200)
        response = self.app.get('/')
        self.assertEqual(response.status_code, 200)

        # 3. Log out
        response = self.app.get('/logout')
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith('/login'))

        # 4. Access index again (should redirect to login)
        response = self.app.get('/')
        self.assertEqual(response.status_code, 302)

    def test_reactive_options_structure(self):
        # 1. Log in
        self.app.post('/login', json={"username": "admin", "password": "admin123"})

        # 2. Query options
        response = self.app.get('/api/options')
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])

        # 3. Verify mapping structure
        branches = data['branches']
        self.assertIn('23/06/2026', branches)
        self.assertIn('26/06/2026', branches)

        # 4. Check branches on specific dates
        damasco_trinidad_branches = branches['23/06/2026'].get('Damasco', [])
        self.assertCountEqual(damasco_trinidad_branches, ['PARAISO - DAMASCO', 'TRINIDAD - DAMASCO'])

        damasco_26_branches = branches['26/06/2026'].get('Damasco', [])
        self.assertCountEqual(damasco_26_branches, ['APURE - DAMASCO', 'AV BOLIVAR VALENCIA - DAMASCO'])

if __name__ == '__main__':
    unittest.main()
