import unittest
import json
import io
from unittest.mock import patch
import pandas as pd
from server import app

# Create mock dataframes for testing
#
# FIX (sesión 25): estos mocks y las aserciones de abajo estaban desactualizados
# respecto al server.py actual (que evolucionó mucho desde que se escribió este
# archivo) y causaban fallos que NO tenían nada que ver con bugs reales:
#   1. Los endpoints /api/* ahora exigen sesión iniciada (is_logged_in()) — sin
#      loguearse primero, todo devolvía 401. Ver setUp().
#   2. df_campo ahora necesita una columna 'Sucursal' unificada (la que arma
#      el script de importación de Kobo vía MATRIZ_SUCURSALES) — antes solo
#      existían las columnas viejas por empresa ('Sede Daka', etc.), y
#      get_options() hace df_date_comp['Sucursal'] directo (sin fallback),
#      así que lanzaba KeyError.
#   3. La fila de Multimax tenía Empresa='dmultimax', pero filtrar_datos_reporte
#      traduce el identificador interno 'dmultimax' a 'mmultimax' al filtrar
#      df_campo (así es como llega el dato crudo real desde Kobo) — con
#      'dmultimax' en el mock, el filtro de empresa nunca matcheaba y el
#      endpoint devolvía 404.
#   4. La hoja del Excel se llama "EDM", no "Dashboard".
#   5. parametros.json ahora configura proyecciones para Multimax también
#      (10%/20%/10%, igual que Damasco) — antes Multimax no tenía proyecciones.
#   6. El valor de 'sucursal' que manda el frontend (static/app.js) es el
#      string tal cual devuelto por /api/options (ej. "DAKA CHACAO"), no el
#      código viejo con guion bajo (ej. "chacao") — usar el código viejo
#      fallaba el match porque el fallback compara guion bajo contra espacio.
MOCK_DF_CAMPO = pd.DataFrame([
    {
        'ID_Envio': 1001,
        'Fecha': '23/06/2026',
        'Horario': '09_00am_10_00am',
        'Empresa': 'ddaka',
        'Sucursal': 'DAKA CHACAO',
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
        'Sucursal': 'SAN DIEGO - DAMASCO',
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
        'Empresa': 'mmultimax',
        'Sucursal': 'MULTIMAX VALENCIA',
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
        # Los endpoints de /api/* exigen sesión iniciada (is_logged_in()) desde
        # que se agregó el sistema de login — sin esto, todos los tests fallaban
        # con 401 aunque fetch_data estuviera mockeado correctamente. Mismo
        # login que usa test_live.py.
        self.app.post('/login', json={"username": "admin", "password": "admin123"})

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

        # 'branches' está indexado primero por fecha y luego por empresa
        # (branches[fecha][empresa] -> [sucursales]), no directo por empresa
        # — así lo arma get_options(). Nombres homologados vía 'Sucursal',
        # que es lo que devuelve /api/options y lo que manda de vuelta el
        # frontend como 'sucursal' (ver static/app.js updateBranches()).
        branches_23_06 = data['branches']['23/06/2026']
        self.assertEqual(branches_23_06['Daka'], ['DAKA CHACAO'])
        self.assertEqual(branches_23_06['Damasco'], ['SAN DIEGO - DAMASCO'])
        self.assertEqual(branches_23_06['Multimax'], ['MULTIMAX VALENCIA'])

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_report_daka(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Daka",
            "sucursal": "DAKA CHACAO",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate',
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        excel_bytes = response.data
        df_dash = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="EDM", header=None)

        # Check Daka projections (20%, 20%, 10%)
        found_t1_ecommerce = False
        found_t2_ecommerce = False
        found_t1_mayor = False
        found_t2_mayor = False
        found_t1_novisibles = False
        found_t2_novisibles = False

        for row in range(df_dash.shape[0]):
            val_col1 = str(df_dash.iloc[row, 1])  # Columna B: etiqueta+% de la Tabla 1
            val_col2 = str(df_dash.iloc[row, 2])  # Columna C: etiqueta (sin %) de la Tabla 2
            val_col4 = str(df_dash.iloc[row, 4])  # Columna E: "+ N%" de la Tabla 2

            if "VENTAS ECOMMERCE/DELIVERY" in val_col1:
                found_t1_ecommerce = True
                self.assertIn("20%", val_col1)
            elif "VENTAS ECOMMERCE/DELIVERY" in val_col2:
                found_t2_ecommerce = True
                self.assertIn("20%", val_col4)

            if "VENTAS AL MAYOR" in val_col1:
                found_t1_mayor = True
                self.assertIn("20%", val_col1)
            elif "VENTAS AL MAYOR" in val_col2:
                found_t2_mayor = True
                self.assertIn("20%", val_col4)

            if "PRODUCTOS NO VISIBLES" in val_col1.upper():
                found_t1_novisibles = True
                self.assertIn("10%", val_col1)
            elif "PRODUCTOS NO VISIBLES" in val_col2.upper():
                found_t2_novisibles = True
                self.assertIn("10%", val_col4)

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
            "sucursal": "SAN DIEGO - DAMASCO",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate',
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        excel_bytes = response.data
        df_dash = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="EDM", header=None)

        # Check Damasco projections (10%, 20%, 10%)
        found_t1_ecommerce = False
        found_t2_ecommerce = False
        found_t1_mayor = False
        found_t2_mayor = False
        found_t1_novisibles = False
        found_t2_novisibles = False

        for row in range(df_dash.shape[0]):
            val_col1 = str(df_dash.iloc[row, 1])
            val_col2 = str(df_dash.iloc[row, 2])
            val_col4 = str(df_dash.iloc[row, 4])

            if "VENTAS ECOMMERCE/DELIVERY" in val_col1:
                found_t1_ecommerce = True
                self.assertIn("10%", val_col1)
            elif "VENTAS ECOMMERCE/DELIVERY" in val_col2:
                found_t2_ecommerce = True
                self.assertIn("10%", val_col4)

            if "VENTAS AL MAYOR" in val_col1:
                found_t1_mayor = True
                self.assertIn("20%", val_col1)
            elif "VENTAS AL MAYOR" in val_col2:
                found_t2_mayor = True
                self.assertIn("20%", val_col4)

            if "PRODUCTOS NO VISIBLES" in val_col1.upper():
                found_t1_novisibles = True
                self.assertIn("10%", val_col1)
            elif "PRODUCTOS NO VISIBLES" in val_col2.upper():
                found_t2_novisibles = True
                self.assertIn("10%", val_col4)

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
            "sucursal": "MULTIMAX VALENCIA",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate',
                                 data=json.dumps(payload),
                                 content_type='application/json')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        excel_bytes = response.data
        df_dash = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="EDM", header=None)

        # Check Multimax projections (10%, 20%, 10%) — parametros.json ya no
        # deja a Multimax sin proyecciones, tiene los mismos % que Damasco.
        found_t1_ecommerce = False
        found_t2_ecommerce = False
        found_t1_mayor = False
        found_t2_mayor = False
        found_t1_novisibles = False
        found_t2_novisibles = False

        for row in range(df_dash.shape[0]):
            val_col1 = str(df_dash.iloc[row, 1])
            val_col2 = str(df_dash.iloc[row, 2])
            val_col4 = str(df_dash.iloc[row, 4])

            if "VENTAS ECOMMERCE/DELIVERY" in val_col1:
                found_t1_ecommerce = True
                self.assertIn("10%", val_col1)
            elif "VENTAS ECOMMERCE/DELIVERY" in val_col2:
                found_t2_ecommerce = True
                self.assertIn("10%", val_col4)

            if "VENTAS AL MAYOR" in val_col1:
                found_t1_mayor = True
                self.assertIn("20%", val_col1)
            elif "VENTAS AL MAYOR" in val_col2:
                found_t2_mayor = True
                self.assertIn("20%", val_col4)

            if "PRODUCTOS NO VISIBLES" in val_col1.upper():
                found_t1_novisibles = True
                self.assertIn("10%", val_col1)
            elif "PRODUCTOS NO VISIBLES" in val_col2.upper():
                found_t2_novisibles = True
                self.assertIn("10%", val_col4)

        self.assertTrue(found_t1_ecommerce)
        self.assertTrue(found_t2_ecommerce)
        self.assertTrue(found_t1_mayor)
        self.assertTrue(found_t2_mayor)
        self.assertTrue(found_t1_novisibles)
        self.assertTrue(found_t2_novisibles)

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_observations_damasco(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Damasco",
            "sucursal": "SAN DIEGO - DAMASCO",
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
        self.assertIn("-HORARIO DE APERTURA Y CIERRE: 9:00AM - 7:00PM", obs_text)
        self.assertIn("-EFECTIVIDAD:", obs_text)

    @patch('server.fetch_data', side_effect=mock_fetch_data)
    def test_generate_observations_multimax(self, mock_fetch):
        payload = {
            "fecha": "23/06/2026",
            "empresa": "Multimax",
            "sucursal": "MULTIMAX VALENCIA",
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
        self.assertIn("-HORARIO DE APERTURA Y CIERRE: 9:00AM - 7:00PM", obs_text)
        self.assertIn("-EFECTIVIDAD:", obs_text)

if __name__ == '__main__':
    unittest.main()
