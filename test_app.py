import unittest
import io
import zipfile
import openpyxl
from app import app

class SummaryAppTestCase(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True
        with self.app.session_transaction() as sess:
            sess['logged_in'] = True
            sess['username'] = 'admin'

    def test_login_and_logout(self):
        client = app.test_client()
        # 1. Unauthenticated request to / redirects to /login
        res_unauth = client.get('/')
        self.assertEqual(res_unauth.status_code, 302)
        self.assertIn('/login', res_unauth.location)

        # 2. Login POST with correct credentials
        res_login = client.post('/login', data={'username': 'info@msgoilfield.com', 'password': 'MSG#2026!SecP@ss982'}, follow_redirects=True)
        self.assertEqual(res_login.status_code, 200)
        self.assertIn(b'MSG Oilfield', res_login.data)

        # 3. Logout GET redirects to /login
        res_logout = client.get('/logout')
        self.assertEqual(res_logout.status_code, 302)
        self.assertIn('/login', res_logout.location)
        response = self.app.get('/')
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'MSG Oilfield', response.data)

    def test_suppliers_api(self):
        response = self.app.get('/api/suppliers')
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])
        self.assertIn("K.HASHIM LLC", data['suppliers'])

    def test_download_internal_stores_excel(self):
        payload = {
            "client": "Dubai Petroleum Establishment",
            "po_number": "50307533",
            "msg_ref": "MSG-1026-2171",
            "is_internal": True,
            "filter_machined_only": True,
            "items": [
                {
                    "sl_no": "1",
                    "client_po_item_no": "10000",
                    "description": "UNION 1\", 6000#",
                    "qty": "3 OF 5",
                    "uom": "pcs",
                    "heat_number": "HN123",
                    "certificate_number": "MTC-001",
                    "make": "WMAASS",
                    "remarks": "Machined from S.40 to Sch.20",
                    "supplier_name": "K.HASHIM LLC",
                    "supplier_po": "PO-1026-3789"
                }
            ]
        }
        response = self.app.post('/api/download', json=payload)
        self.assertEqual(response.status_code, 200)
        
        excel_bytes = io.BytesIO(response.data)
        wb = openpyxl.load_workbook(excel_bytes)
        ws = wb.active
        self.assertEqual(ws['A1'].value, "MSG OILFIELD - STORES SUMMARY SHEET")
        self.assertEqual(ws['A5'].value, "SL NO")
        self.assertEqual(ws['B5'].value, "CLIENT PO ITEM NO")
        self.assertEqual(ws['C5'].value, "Description")
        self.assertEqual(ws['B6'].value, "10000")

    def test_export_supplier_pdfs_zip(self):
        payload = {
            "client": "Dubai Petroleum Establishment",
            "po_number": "50307533",
            "msg_ref": "MSG-1026-2171",
            "items": [
                {
                    "sl_no": "1",
                    "client_po_item_no": "10000",
                    "description": "UNION 1\", 6000#",
                    "qty": "12.00",
                    "uom": "pcs",
                    "heat_number": "HN999",
                    "certificate_number": "CERT-111",
                    "make": "WMAASS",
                    "remarks": "Locally made from plate",
                    "supplier_name": "K.HASHIM LLC",
                    "supplier_po": "PO-1026-3789"
                }
            ]
        }
        response = self.app.post('/api/export-supplier-pdfs', json=payload)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, 'application/zip')

    def test_stock_supplier_and_offered_remark_parsing(self):
        from pdf_parser import strip_internal_product_code, strip_leading_item_number, auto_detect_make, REMARK_PATTERN
        from excel_generator import is_meaningful_item

        # 1. Test code with STOCK
        desc = "BL25150RFA105/A350LF2ST&H-STOCK FLANGE 2\", 150#, BLIND RF"
        cleaned = strip_internal_product_code(desc)
        self.assertEqual(cleaned, "FLANGE 2\", 150#, BLIND RF")

        # 2. Test leading item number stripping (e.g. '1 ECC. RED...' -> 'ECC. RED...')
        self.assertEqual(strip_leading_item_number("1 ECC. RED. 6\" X 3\"", sl_no=1), "ECC. RED. 6\" X 3\"")
        self.assertEqual(strip_leading_item_number("2 ECC. RED. 8\" X 6\"", sl_no=2), "ECC. RED. 8\" X 6\"")
        self.assertEqual(strip_leading_item_number("1\" FLANGE 150#", sl_no=1), "1\" FLANGE 150#")

        # 3. Test MAKE auto-detection mappings
        self.assertEqual(auto_detect_make("FLANGE 2\" MFF A105N"), "Metalfar Italy")
        self.assertEqual(auto_detect_make("TEE 3\" BENKAN A234 WPB"), "BENKAN THAILAND")
        self.assertEqual(auto_detect_make("BL25150RFA105/A350LF2ST&H-STOCK"), "ST&H KOREA")
        self.assertEqual(auto_detect_make("XEL90250S40SRWPB-W-STOCK"), "WMASS GERMANY")
        self.assertEqual(auto_detect_make("EL90256KNPTA105JD-STOCK"), "DELCORTE GERMANY")
        self.assertEqual(auto_detect_make("LTO156KA350LF2-OMSA-STOCK"), "OMSA GERMANY")
        self.assertEqual(auto_detect_make("XEL90250S40SRWPB-TK-STOCK"), "TK KOREA")
        self.assertEqual(auto_detect_make("SPECTACLE BLIND 3/4\" ULMA 300#"), "ULMA SPAIN")
        self.assertEqual(auto_detect_make("FLANGE 6\" MELESI 300# BLIND"), "MELESI ITALY")

        # 4. Test Offered remark matching
        item_offered = {
            "description": "UNION 1\", 6000#",
            "remarks": "Offered 1/2\" NPT connection"
        }
        self.assertTrue(is_meaningful_item(item_offered))

        # 5. Test remark pattern matching
        m = REMARK_PATTERN.search("SPECTACLE BLIND 3/4\" - Offered 1/2\" NPT")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(0), "Offered 1/2\" NPT")

    def test_history_and_cloudinary_api(self):
        # 1. Test Cloudinary Config API GET & POST
        res = self.app.get('/api/cloudinary-config')
        self.assertEqual(res.status_code, 200)
        
        # 2. Test Save Worksheet API
        payload = {
            "client": "Test Client LLC",
            "po_number": "PO-9999",
            "msg_ref": "MSG-TEST-001",
            "items": [{"sl_no": 1, "description": "FLANGE 2\" 150#"}]
        }
        res_save = self.app.post('/api/save-worksheet', json=payload)
        self.assertEqual(res_save.status_code, 200)
        data_save = res_save.get_json()
        self.assertTrue(data_save['success'])
        sheet_id = data_save['record']['id']

        # 3. Test History GET API (with search query)
        res_hist = self.app.get('/api/history?q=Test Client')
        self.assertEqual(res_hist.status_code, 200)
        sheets = res_hist.get_json()['sheets']
        self.assertTrue(any(s['id'] == sheet_id for s in sheets))

        # 4. Test Get History Detail API
        res_detail = self.app.get(f'/api/history/{sheet_id}')
        self.assertEqual(res_detail.status_code, 200)
        self.assertEqual(res_detail.get_json()['sheet']['client'], "Test Client LLC")

        # 5. Test Delete History API
        res_del = self.app.delete(f'/api/history/{sheet_id}')
        self.assertEqual(res_del.status_code, 200)

if __name__ == '__main__':
    unittest.main()
