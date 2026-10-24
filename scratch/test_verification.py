import os
import sys
import shutil

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app

def run_tests():
    print("=== STARTING FULL VERIFICATION ===")
    client = app.app.test_client()

    # 1. Test index page
    res = client.get('/')
    assert res.status_code == 200, f"Index failed: {res.status_code}"
    assert b"Office Automation" in res.data, "Brand missing in index"
    assert b"mis-active-filename" in res.data, "Active filename pill missing in HTML"
    assert b"chk-set-master-excel" in res.data, "Set this as master excel checkbox missing in HTML"
    assert b"estimate-active-filename" in res.data, "Estimate active filename pill missing in HTML"
    assert b"chk-set-estimate-excel" in res.data, "Set as default estimate excel checkbox missing in HTML"
    assert b"mis-excel-dropdown" not in res.data, "Old mis-excel-dropdown still found in HTML!"
    assert b"web-excel-dropdown" not in res.data, "Old web-excel-dropdown still found in HTML!"
    assert b"estimate-excel-dropdown" not in res.data, "Old estimate-excel-dropdown still found in HTML!"
    print("PASS: Index HTML verified: old scanning dropdowns removed, active pills & master checkboxes present.")

    # 1b. Test list-local-excels returns immediately with empty list (no PC scan)
    res = client.get('/api/list-local-excels')
    assert res.status_code == 200
    assert res.json == {"success": True, "files": []}
    print("PASS: /api/list-local-excels verified as instant no-op (no disk scanning).")

    # 2. Test config endpoint
    res = client.get('/api/config')
    assert res.status_code == 200
    cfg_data = res.json
    print("PASS: /api/config returned valid JSON configuration.")

    # 3. Test excel file status
    res = client.get('/api/excel-file-status?type=master')
    assert res.status_code == 200
    master_status = res.json
    assert master_status['success'] is True
    assert master_status['exists'] is True
    print(f"PASS: /api/excel-file-status (master): {master_status['filename']}, mtime: {master_status['last_modified']}")

    res = client.get('/api/excel-file-status?type=estimate')
    assert res.status_code == 200
    est_status = res.json
    assert est_status['success'] is True
    assert est_status['exists'] is True
    print(f"PASS: /api/excel-file-status (estimate): {est_status['filename']}, mtime: {est_status['last_modified']}")

    # 4. Test set-excel-path with save_as_master
    master_path = cfg_data['config']['master_excel_path']
    res = client.post('/api/set-excel-path', json={'file_path': master_path, 'type': 'master', 'save_as_master': True})
    assert res.status_code == 200
    data = res.json
    assert data['success'] is True
    assert len(data['sheets']) > 0
    assert data['next_ref_no'] is not None
    print(f"PASS: /api/set-excel-path (master): Linked {data['filename']}, next_ref: {data['next_ref_no']}")

    # 5. Test append with a test copy so we don't mess up original data
    test_excel = os.path.join(os.path.dirname(master_path), "test_append_copy.xlsx")
    shutil.copy2(master_path, test_excel)
    try:
        # Set config to test copy
        cfg = app.load_config()
        cfg["master_excel_path"] = test_excel
        app.save_config(cfg)

        test_row = {
            "ref_no": 9999,
            "ins_co": "TEST_CO",
            "policy_no": "POL-12345",
            "claim_no": "CLM-67890",
            "insured_name": "Test User",
            "vehicle_no": "DL01AB1234",
            "make_model": "HYUNDAI CRETA",
            "workshop": "TEST WORKSHOP",
            "date_of_survey": "24-Oct-2026"
        }
        res = client.post('/api/append', json={'rows': [test_row]})
        assert res.status_code == 200
        append_res = res.json
        assert append_res['success'] is True
        assert 'last_modified' in append_res
        assert 'file_size' in append_res
        assert append_res['inserted_count'] == 1
        print(f"PASS: /api/append verified! Inserted row {append_res['start_row']}, mtime: {append_res['last_modified']}")
    finally:
        # Restore original master path and remove test copy
        cfg["master_excel_path"] = master_path
        app.save_config(cfg)
        if os.path.exists(test_excel):
            os.remove(test_excel)
        # Also clean up backup of test file if created
        for f in os.listdir(os.path.dirname(master_path)):
            if f.startswith("test_append_copy.xlsx.bak_"):
                try:
                    os.remove(os.path.join(os.path.dirname(master_path), f))
                except Exception:
                    pass

    # 6. Test write-estimate with a test copy of the estimate excel
    est_path = cfg_data['config'].get('estimate_excel_path')
    if est_path and os.path.exists(est_path):
        est_ext = os.path.splitext(est_path)[1]
        test_est_copy = os.path.join(os.path.dirname(est_path), f"test_est_copy{est_ext}")
        shutil.copy2(est_path, test_est_copy)
        try:
            sample_part = {
                "description": "FRONT BUMPER COVER ASSY",
                "amount": 15450.0,
                "selected": True,
                "item_type": "Part"
            }
            res = client.post('/api/write-estimate', json={
                'file_path': test_est_copy,
                'sheet_name': '',
                'parts': [sample_part],
                'start_row': 171,
                'bill_no': 'TEST-INV-001'
            })
            assert res.status_code == 200
            est_write_res = res.json
            assert est_write_res['success'] is True
            assert est_write_res['written_count'] == 1
            assert est_write_res['start_row'] == 171
            print(f"PASS: /api/write-estimate verified! Written row {est_write_res['start_row']} to {est_write_res['end_row']}, mtime: {est_write_res['last_modified']}")
        finally:
            cfg = app.load_config()
            cfg["estimate_excel_path"] = est_path
            app.save_config(cfg)
            if os.path.exists(test_est_copy):
                os.remove(test_est_copy)
            for f in os.listdir(os.path.dirname(est_path)):
                if f.startswith(f"test_est_copy{est_ext}.bak_"):
                    try:
                        os.remove(os.path.join(os.path.dirname(est_path), f))
                    except Exception:
                        pass

    # 7. Check Progress Bar UI Elements in index.html
    res = client.get('/')
    assert b"progress-modal-overlay" in res.data, "progress-modal-overlay missing from index.html"
    assert b"progress-bar-fill" in res.data, "progress-bar-fill missing from index.html"
    assert b"showProgress" in res.data, "showProgress JS function missing from index.html"
    print("PASS: Progress bar HTML/CSS/JS components successfully verified in index.html.")

    print("=== ALL VERIFICATION TESTS PASSED SUCCESSFULLY! ===")

if __name__ == "__main__":
    run_tests()
