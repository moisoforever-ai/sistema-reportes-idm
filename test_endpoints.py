import unittest
import json
import io
from unittest.mock import patch
import pandas as pd
from server import app

# Create mock dataframes for testing
MOCK_DF_CAMPO = pd.DataFrame([
    {
        'ID_Envio': 1001,
        'Fecha': '23/06/2026',
        'Horario': '09_00am_10_00am',
        'Empresa': 'ddaka',
        'Sede Daka': 'chacao',
        'Sedes Damasco': None,
        'Sedes Multimax': None,
        'Producto': 'television 32',
        'Marca': 'samsung',
        'Cantidad': 2,
        'Precio': 0,
        'Factura': 5001.0,
        'Visitas': 20.0
    },
    {
        'ID_Envio': 1002,
        'Fecha': '23/06/2026',
        'Horario': '10_00am_11_00am',
        'Empresa': 'ddamasco',
        'Sede Daka': None,
        'Sedes Damasco': 'san_diego',
        'Sedes Multimax': None,
        'Producto': 'cocina a gas 4 hornillas',
        'Marca': 'da+co',
        'Cantidad': 1,
        'Precio': 0,
        'Factura': 5002.0,
        'Visitas': 15.0
    },
    {
        'ID_Envio': 1003,
        'Fecha': '23/06/2026',
        'Horario': '11_00am_12_00pm',
        'Empresa': 'dmultimax',
        'Sede Daka': None,
        'Sedes Damasco': None,
        'Sedes Multimax': 'valencia',
        'Producto': 'licuadora oster',
        'Marca': 'oster',
        'Cantidad': 3,
        'Precio': 0,
        'Factura': 5003.0,
        'Visitas': 30.0
    }
])

MOCK_MAESTRO_DAKA = pd.DataFrame([
    {
        'PRODUCTO': 'television 32',
        'MARCA': 'SAMSUNG',
        'PRECIO DE REFERENCIA': '150'
    }
])

MOCK_MAESTRO_DAMASCO = pd.DataFrame([
    {
        'PRODUCTO': 'cocina a gas 4 hornillas',
        'MARCA': 'DA+CO',
        'PRECIO DE REFERENCIA': '250'
    }
])

MOCK_MAESTRO_MULTIMAX = pd.DataFrame([
    {
        'PRODUCTO': 'licuadora oster',
        'MARCA': 'OSTER',
        'PRECIO DE REFERENCIA': '60'
    }
])

def mock_fetch_data(gid_base=None):
    # Returns (df_campo, df_maestro)
    # Depending on GID, return the correct maestro
    from server import GID_BASE_DAKA, GID_BASE_DAMASCO, GID_BASE_MULTIMAX
    if gid_base == GID_BASE_DAMASCO:
        return MOCK_DF_CAMPO, MOCK_MAESTRO_DAMASCO
    elif gid_base == GID_BASE_MULTIMAX:
        return MOCK_DF_CAMPO, MOCK_MAESTRO_MULTIMAX
    else:
        return MOCK_DF_CAMPO, MOCK_MAESTRO_DAKA

class TestServerEndpoints(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_options(self, mock_fetch):
        response = self.app.get('/api/options')
        self.assertEqual(response.status_code, 200)
        data = json.loads(response.data.decode('utf-8'))
        self.assertTrue(data['success'])
        
        print("\n--- [MOCKED] API OPTIONS ---")
        print("Dates available:", data['dates'])
        print("Companies available:", data['companies'])
        print("Branches mapping:")
        for comp, branches in data['branches'].items():
            print(f"  {comp}: {branches}")

        self.assertIn('Daka', data['companies'])
        self.assertIn('Damasco', data['companies'])
        self.assertIn('Multimax', data['companies'])
        
        # Verify branches map correctly
        self.assertEqual(data['branches']['Daka'], ['chacao'])
        self.assertEqual(data['branches']['Damasco'], ['san_diego'])
        self.assertEqual(data['branches']['Multimax'], ['valencia'])

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_report_daka(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Daka",
            "sucursal": "chacao",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate', 
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        excel_bytes = response.data
        df_dash = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Dashboard", header=None)
        
        # Check Daka projections (20%, 20%, 10%)
        found_t1_ecommerce = False
        found_t2_ecommerce = False
        found_t1_mayor = False
        found_t2_mayor = False
        found_t1_novisibles = False
        found_t2_novisibles = False
        
        for row in range(df_dash.shape[0]):
            val_col0 = str(df_dash.iloc[row, 0])
            val_col2 = str(df_dash.iloc[row, 2])
            
            if "VENTAS ECOMMERCE/DELIVERY" in val_col0:
                if "(" in val_col0:
                    found_t1_ecommerce = True
                    self.assertIn("20%", val_col0)
                else:
                    found_t2_ecommerce = True
                    self.assertIn("20%", val_col2)
            if "VENTAS AL MAYOR" in val_col0:
                if "(" in val_col0:
                    found_t1_mayor = True
                    self.assertIn("20%", val_col0)
                else:
                    found_t2_mayor = True
                    self.assertIn("20%", val_col2)
            if "PRODUCTOS NO VISIBLES" in val_col0 or "PRODUCTOS NO VISIBLES" in val_col0.upper():
                if "(" in val_col0:
                    found_t1_novisibles = True
                    self.assertIn("10%", val_col0)
                else:
                    found_t2_novisibles = True
                    self.assertIn("10%", val_col2)
                
        self.assertTrue(found_t1_ecommerce)
        self.assertTrue(found_t2_ecommerce)
        self.assertTrue(found_t1_mayor)
        self.assertTrue(found_t2_mayor)
        self.assertTrue(found_t1_novisibles)
        self.assertTrue(found_t2_novisibles)

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_report_damasco(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Damasco",
            "sucursal": "san_diego",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate', 
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        excel_bytes = response.data
        df_dash = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Dashboard", header=None)
        
        # Check Damasco projections (10%, 20%, 10%)
        found_t1_ecommerce = False
        found_t2_ecommerce = False
        found_t1_mayor = False
        found_t2_mayor = False
        found_t1_novisibles = False
        found_t2_novisibles = False
        
        for row in range(df_dash.shape[0]):
            val_col0 = str(df_dash.iloc[row, 0])
            val_col2 = str(df_dash.iloc[row, 2])
            
            if "VENTAS ECOMMERCE/DELIVERY" in val_col0:
                if "(" in val_col0:
                    found_t1_ecommerce = True
                    self.assertIn("10%", val_col0)
                else:
                    found_t2_ecommerce = True
                    self.assertIn("10%", val_col2)
            if "VENTAS AL MAYOR" in val_col0:
                if "(" in val_col0:
                    found_t1_mayor = True
                    self.assertIn("20%", val_col0)
                else:
                    found_t2_mayor = True
                    self.assertIn("20%", val_col2)
            if "PRODUCTOS NO VISIBLES" in val_col0 or "PRODUCTOS NO VISIBLES" in val_col0.upper():
                if "(" in val_col0:
                    found_t1_novisibles = True
                    self.assertIn("10%", val_col0)
                else:
                    found_t2_novisibles = True
                    self.assertIn("10%", val_col2)
                
        self.assertTrue(found_t1_ecommerce)
        self.assertTrue(found_t2_ecommerce)
        self.assertTrue(found_t1_mayor)
        self.assertTrue(found_t2_mayor)
        self.assertTrue(found_t1_novisibles)
        self.assertTrue(found_t2_novisibles)

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_report_multimax(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Multimax",
            "sucursal": "valencia",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate', 
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        excel_bytes = response.data
        df_dash = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Dashboard", header=None)
        
        # Check that Multimax has NO projections (no Ecommerce/Mayor/No Visibles labels)
        for row in range(df_dash.shape[0]):
            val = str(df_dash.iloc[row, 0])
            self.assertNotIn("VENTAS ECOMMERCE/DELIVERY", val)
            self.assertNotIn("VENTAS AL MAYOR", val)
            self.assertNotIn("PRODUCTOS NO VISIBLES", val)

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_observations_damasco(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Damasco",
            "sucursal": "san_diego",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate_observations', 
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "text/plain")
        obs_text = response.data.decode('utf-8')
        
        print("\n--- [MOCKED] GENERATE OBSERVATIONS DAMASCO ---")
        print(obs_text)
        self.assertIn("-Horario de apertura y cierre: 9:00AM - 7:00PM", obs_text)
        self.assertIn("-Efectividad:", obs_text)

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_observations_multimax(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Multimax",
            "sucursal": "valencia",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate_observations', 
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "text/plain")
        obs_text = response.data.decode('utf-8')
        
        print("\n--- [MOCKED] GENERATE OBSERVATIONS MULTIMAX ---")
        print(obs_text)
        self.assertIn("-Horario de apertura y cierre: 9:00AM - 7:00PM", obs_text)

if __name__ == '__main__':
    unittest.main()
