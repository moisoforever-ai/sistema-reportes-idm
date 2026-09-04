import unittest
import os
import json
from server import app, USERS_FILE, load_users

class TestAdminUserMgmt(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True
        
        # Reset users.json to default state before each test
        default_users = {
            "admin": "admin123",
            "idm": "idm2026"
        }
        with open(USERS_FILE, 'w', encoding='utf-8') as f:
            json.dump(default_users, f, indent=4)

    def tearDown(self):
        # Clean up users.json after tests
        if os.path.exists(USERS_FILE):
            try:
                os.remove(USERS_FILE)
            except Exception:
                pass

    def test_non_admin_cannot_access_user_mgmt(self):
        # Log in as idm (non-admin)
        self.app.post('/login', json={"username": "idm", "password": "idm2026"})
        
        # Try to access admin panel (should redirect to home /)
        response = self.app.get('/admin/users')
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith(('/', '/home')))

    def test_admin_can_access_user_mgmt(self):
        # Log in as admin
        self.app.post('/login', json={"username": "admin", "password": "admin123"})
        
        # Try to access admin panel (should be 200)
        response = self.app.get('/admin/users')
        self.assertEqual(response.status_code, 200)

    def test_admin_add_and_delete_user(self):
        # 1. Log in as admin
        self.app.post('/login', json={"username": "admin", "password": "admin123"})
        
        # 2. Add user "newuser"
        response = self.app.post('/admin/users', data={
            "action": "add",
            "new_username": "newuser",
            "new_password": "newpassword"
        }, follow_redirects=True)
        
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"newuser", response.data)
        
        # 3. Verify user is saved in users.json
        users = load_users()
        self.assertIn("newuser", users)
        self.assertEqual(users["newuser"], "newpassword")
        
        # 4. Log out and try logging in as newuser
        self.app.get('/logout')
        response = self.app.post('/login', json={"username": "newuser", "password": "newpassword"})
        self.assertEqual(response.status_code, 200)
        
        # 5. Log back in as admin and delete the user
        self.app.get('/logout')
        self.app.post('/login', json={"username": "admin", "password": "admin123"})
        
        response = self.app.post('/admin/users', data={
            "action": "delete",
            "delete_username": "newuser"
        }, follow_redirects=True)
        
        self.assertEqual(response.status_code, 200)
        
        # 6. Verify user is removed from users.json
        users = load_users()
        self.assertNotIn("newuser", users)

    def test_cannot_delete_admin(self):
        # Log in as admin
        self.app.post('/login', json={"username": "admin", "password": "admin123"})
        
        # Try to delete admin
        response = self.app.post('/admin/users', data={
            "action": "delete",
            "delete_username": "admin"
        }, follow_redirects=True)
        
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"No se puede eliminar el usuario", response.data)
        
        # Admin should still be there
        users = load_users()
        self.assertIn("admin", users)

if __name__ == '__main__':
    unittest.main()
