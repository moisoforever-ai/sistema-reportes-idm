import unittest
import io
import pandas as pd
from server import app

class TestLiveEndpoints(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True
        # Log in to authenticate test requests
        self.app.post('/login', json={"username": "admin", "password": "admin123"})

    def test_live_options(self):
        print("\n--- Testing Live /api/options ---")
        response = self.app.get('/api/options')
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])
        print("Available dates:", data['dates'])
        print("Available companies:", data['companies'])
        print("Damasco branches:", data['branches'].get('Damasco', []))

    def test_live_generate_report_damasco_trinidad(self):
        print("\n--- Testing Live /api/generate (Damasco CC Recreo 30/06/2026) ---")
        payload = {
            "fecha": "30/06/2026",
            "empresa": "Damasco",
            "sucursal": "BARQUISIMETO CC RECREO - DAMASCO",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate', json=payload)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # Verify columns and details
        excel_bytes = response.data
        df_detailed = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Datos Detallados")
        df_dash = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="EDM", header=None)
        
        print("Detailed records found:", len(df_detailed))
        print("Standardized products:")
        print(df_detailed[['PRODUCTO CAMPO', 'PRODUCTO ESTANDARIZADO', 'PRECIO ($)', 'CANTIDAD', 'VENTA TOTAL ($)', 'ES PROMO', 'MARCA PROPIA']].to_string())
        
        # Verify that all products belong to the Damasco Master DB
        # E.g., no "Hyundai" products should be matched unless they are part of the Damasco Master DB (which has none).
        for idx, row in df_detailed.iterrows():
            prod_std = str(row['PRODUCTO ESTANDARIZADO'])
            self.assertNotIn("Hyundai", prod_std)
            self.assertNotIn("Bremen", prod_std)
            self.assertNotIn("Nasa", prod_std)
            # Check laptop is matched to a laptop product, not a morral
            if "laptop" in str(row['PRODUCTO CAMPO']).lower():
                self.assertIn("Laptop", prod_std)
                self.assertNotIn("Morral", prod_std)
        
        # Check projection values for Damasco (10%, 20%, 10%)
        found_t1_ecommerce = False
        found_t2_ecommerce = False
        found_t1_mayor = False
        found_t2_mayor = False
        found_t1_novisibles = False
        found_t2_novisibles = False
        
        for row in range(df_dash.shape[0]):
            val_col1 = str(df_dash.iloc[row, 1]) # Column B
            val_col2 = str(df_dash.iloc[row, 2]) # Column C
            val_col4 = str(df_dash.iloc[row, 4]) # Column E
            
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
                
            if "PRODUCTOS NO VISIBLES" in val_col1 or "PRODUCTOS NO VISIBLES" in val_col1.upper():
                found_t1_novisibles = True
                self.assertIn("10%", val_col1)
            elif "PRODUCTOS NO VISIBLES" in val_col2 or "PRODUCTOS NO VISIBLES" in val_col2.upper():
                found_t2_novisibles = True
                self.assertIn("10%", val_col4)
                
        self.assertTrue(found_t1_ecommerce)
        self.assertTrue(found_t2_ecommerce)
        self.assertTrue(found_t1_mayor)
        self.assertTrue(found_t2_mayor)
        self.assertTrue(found_t1_novisibles)
        self.assertTrue(found_t2_novisibles)
        print("Damasco projections verified in Excel.")

        # Check Large Sale Discount
        # In our script, the only large invoice was ID_Envio: 788889160 | Factura: 3.0 (sum: $1020)
        # Let's locate the Large Sale Discount row in the EDM sheet
        large_discount_val = None
        for row in range(df_dash.shape[0]):
            val_col2 = str(df_dash.iloc[row, 2])
            if "DESCUENTO EN PROYECCION POR VENTA GRANDE" in val_col2:
                large_discount_val = df_dash.iloc[row, 5] # Column F (index 5)
                break
        
        print("Large Sale Discount in EDM:", large_discount_val)
        self.assertEqual(large_discount_val, 0)

    def test_live_generate_observations_damasco_trinidad(self):
        print("\n--- Testing Live /api/generate_observations (Damasco CC Recreo 30/06/2026) ---")
        payload = {
            "fecha": "30/06/2026",
            "empresa": "Damasco",
            "sucursal": "BARQUISIMETO CC RECREO - DAMASCO",
            "hora_apertura": "09:00",
            "hora_cierre": "19:00"
        }
        response = self.app.post('/api/generate_observations', json=payload)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "text/plain")
        obs_text = response.data.decode('utf-8')
        print("Observations Output:")
        print(obs_text)
        self.assertIn("-HORARIO DE APERTURA Y CIERRE: 9:00AM - 7:00PM", obs_text)
        self.assertIn("-EFECTIVIDAD:", obs_text)
        self.assertIn("no se observo descuento en proyeccion por venta grande", obs_text.lower())

if __name__ == '__main__':
    unittest.main()
