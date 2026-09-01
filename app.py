from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, send_file
import mysql.connector
from datetime import date
from flask import g  # to prevent database timeout
import openpyxl
import io
import csv
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = 'tribi_secret_key'

_grn_schema_checked = False

def ensure_grn_schema():
    global _grn_schema_checked
    if _grn_schema_checked:
        return
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="admin",
            database="tribi_db"
        )
        cursor = conn.cursor(dictionary=True)
        
        cursor.execute("SHOW COLUMNS FROM grn")
        grn_cols = [r['Field'] for r in cursor.fetchall()]
        if 'po_id' not in grn_cols:
            cursor.execute("ALTER TABLE grn ADD COLUMN po_id INT DEFAULT NULL")
            conn.commit()
        if 'po_number' not in grn_cols:
            cursor.execute("ALTER TABLE grn ADD COLUMN po_number VARCHAR(50) DEFAULT NULL")
            conn.commit()
            
        cursor.execute("SHOW COLUMNS FROM grn_item")
        item_cols = [r['Field'] for r in cursor.fetchall()]
        if 'posted_qty_accepted' not in item_cols:
            cursor.execute("ALTER TABLE grn_item ADD COLUMN posted_qty_accepted DECIMAL(10,3) NOT NULL DEFAULT 0.000")
            cursor.execute("UPDATE grn_item SET posted_qty_accepted = qty_accepted WHERE qty_accepted IS NOT NULL")
            conn.commit()

        # Deduplicate storage table rows per item_id and add UNIQUE KEY
        cursor.execute("SELECT item_id, COUNT(*) as cnt, SUM(physical_availability) as total_qty, MIN(store_id) as keep_store_id FROM storage GROUP BY item_id HAVING cnt > 1")
        dup_items = cursor.fetchall()
        for dup in dup_items:
            cursor.execute("UPDATE storage SET physical_availability = %s WHERE store_id = %s", (dup['total_qty'], dup['keep_store_id']))
            cursor.execute("DELETE FROM storage WHERE item_id = %s AND store_id != %s", (dup['item_id'], dup['keep_store_id']))
        if dup_items:
            conn.commit()

        cursor.execute("SHOW INDEX FROM storage WHERE Key_name = 'uq_storage_item'")
        if not cursor.fetchall():
            cursor.execute("ALTER TABLE storage ADD UNIQUE KEY uq_storage_item (item_id)")
            conn.commit()
            
        cursor.close()
        conn.close()
        _grn_schema_checked = True
    except Exception as e:
        print("Error ensuring GRN/Storage schema:", e)

def get_db():
    ensure_grn_schema()
    # to prevent database timeout
    conn = mysql.connector.connect(
        host="localhost",
        user="root",
        password="admin",
        database="tribi_db"
    )
    # to prevent database timeout
    setup_cursor = conn.cursor()
    setup_cursor.execute("SET SESSION innodb_lock_wait_timeout = 10")
    setup_cursor.close()

    # to prevent database timeout: track this connection on the request context so it can be force-closed
    g.db_connections = getattr(g, 'db_connections', [])
    g.db_connections.append(conn)

    return conn

STATUS_MAP = {
    0: ('Open', 'badge-open'),
    1: ('In Progress', 'badge-inprogress'),
    2: ('Active', 'badge-active'),
    3: ('Closed', 'badge-closed'),
    4: ('Cancelled', 'badge-cancelled'),
    5: ('Reopened', 'badge-reopened')
}

GRN_TYPE_MAP = {
    0: 'Purchase',
    1: 'Back To Store',
    2: 'Internal Return',
    3: 'Online Purchase',
    4: 'Any Type',
    5: 'Work Order Return'
}

# to prevent database timeout
@app.teardown_request
def cleanup_db_on_error(exception=None):
    """Runs after EVERY request, success or failure. If the request crashed
    with an exception, any MySQL connections it opened via get_db() are rolled
    back and closed here — so their locks are released immediately instead of
    sitting open until MySQL's timeout kicks in and blocks the next user's
    request. On the normal success path your routes already commit/close
    themselves, so this has nothing to do."""
    if exception is not None:
        for conn in getattr(g, 'db_connections', []):
            try:
                conn.rollback()
            except Exception:
                pass
            try:
                conn.close()
            except Exception:
                pass

# to prevent database timeout
@app.errorhandler(mysql.connector.errors.DatabaseError)
def handle_db_lock_timeout(e):
    if getattr(e, 'errno', None) == 1205:
        flash('This record is being updated by someone else right now. Please wait a moment and try again.', 'error')
        return redirect(request.referrer or url_for('dashboard'))
    raise e

            
def get_setting(key, default=None, cursor=None):
    close_cursor = False
    if cursor is None:
        try:
            db = get_db()
            cursor = db.cursor(dictionary=True)
            close_cursor = True
        except Exception:
            return default
    try:
        cursor.execute("SELECT setting_value FROM system_settings WHERE setting_key = %s", (key,))
        row = cursor.fetchone()
        val = row['setting_value'] if row else default
    except Exception:
        val = default
    finally:
        if close_cursor and cursor:
            try:
                cursor.close()
                db.close()
            except Exception:
                pass
    return val


def set_setting(key, value, cursor=None):
    close_cursor = False
    if cursor is None:
        try:
            db = get_db()
            cursor = db.cursor(dictionary=True)
            close_cursor = True
        except Exception:
            return
    try:
        cursor.execute("""
            INSERT INTO system_settings (setting_key, setting_value)
            VALUES (%s, %s)
            ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value)
        """, (key, str(value)))
        if close_cursor:
            db.commit()
    except Exception as e:
        print(f"Error setting {key}:", e)
    finally:
        if close_cursor and cursor:
            try:
                cursor.close()
                db.close()
            except Exception:
                pass


def get_financial_year(cursor=None, date_obj=None):
    custom_enabled = get_setting('fy_custom_enabled', '0', cursor=cursor)
    if custom_enabled == '1':
        fy_label = get_setting('fy_label', '', cursor=cursor)
        if fy_label:
            return fy_label
        start_year = get_setting('fy_start_year', '', cursor=cursor)
        end_year = get_setting('fy_end_year', '', cursor=cursor)
        if start_year and end_year:
            if start_year == end_year:
                return str(start_year)
            return f"{start_year}-{end_year}"
        elif start_year:
            return str(start_year)

    if date_obj is None:
        dt = date.today()
    elif isinstance(date_obj, str):
        try:
            dt = datetime.strptime(date_obj, '%Y-%m-%d').date()
        except Exception:
            dt = date.today()
    else:
        dt = date_obj

    if dt.month >= 4:
        return f"{dt.year}-{dt.year + 1}"
    else:
        return f"{dt.year - 1}-{dt.year}"


def get_financial_year_start(cursor=None, date_obj=None):
    custom_enabled = get_setting('fy_custom_enabled', '0', cursor=cursor)
    if custom_enabled == '1':
        start_year = get_setting('fy_start_year', '', cursor=cursor)
        if start_year:
            return str(start_year)
    fy = get_financial_year(cursor=cursor, date_obj=date_obj)
    if '-' in fy:
        return fy.split('-')[0]
    return fy


def generate_wo_number(cursor):
    fy = get_financial_year(cursor=cursor)
    cursor.execute("""SELECT wo_number FROM work_order 
                      WHERE wo_number LIKE %s 
                      ORDER BY wo_id DESC LIMIT 1""", (f'%/{fy}',))
    last = cursor.fetchone()
    if last and last.get('wo_number'):
        try:
            seq = int(last['wo_number'].split('/')[0]) + 1
        except (ValueError, IndexError):
            seq = 2537
    else:
        seq = 2537
    return f"{seq}/{fy}"


def generate_grn_number(cursor):
    custom_enabled = get_setting('fy_custom_enabled', '0', cursor=cursor)
    grn_format = get_setting('fy_grn_format', 'start_year', cursor=cursor)
    if custom_enabled == '1':
        if grn_format == 'full_range':
            year_str = get_financial_year(cursor=cursor)
        else:
            year_str = get_financial_year_start(cursor=cursor)
    else:
        year_str = str(date.today().year)

    prefix = f"GRN-{year_str}-"
    cursor.execute(
        "SELECT grn_no FROM grn WHERE grn_no LIKE %s ORDER BY grn_id DESC LIMIT 1",
        (f"{prefix}%",)
    )
    last = cursor.fetchone()
    if last and last.get('grn_no'):
        try:
            seq = int(last['grn_no'].split('-')[-1]) + 1
        except (ValueError, IndexError):
            seq = 1
    else:
        seq = 1
    return f"{prefix}{seq:04d}"


def get_bom_from_sheets(sheetnames):
    for name in sheetnames:
        cleaned = name.strip()
        if cleaned.lower().startswith('purchase'):
            # The suffix should be the characters after 'purchase' (8 characters)
            suffix = cleaned[8:]
            # Strip any leading separators like '_', '-', or spaces
            bom_code = suffix.lstrip('_ -')
            if bom_code:
                return bom_code, name
    return None, None


# ─── PURCHASE SHEET HELPERS ──────────────────────────────
def find_purchase_sheet_name(workbook):
    """Find the sheet whose name identifies it as the purchase sheet
    (e.g. 'Purchase-300044124028'), ignoring all the other sheets in the
    workbook. Falls back sensibly if no such sheet is found."""
    for name in workbook.sheetnames:
        if name.strip().lower().startswith('purchase'):
            return name
    for name in workbook.sheetnames:
        if 'purchase' in name.strip().lower():
            return name
    return workbook.sheetnames[0]


def parse_purchase_sheet(file_stream):
    """Reads only the 'Purchase-...' sheet out of a multi-sheet workbook and
    returns a list of dicts: item_code, qty, mpn, manufacturer_name, desc."""
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    sheet_name = find_purchase_sheet_name(wb)
    ws = wb[sheet_name]

    def clean(v):
        return str(v).strip().lower() if v is not None else ''

    # Locate the header row - scan the first 10 rows for one that looks
    # like a header (mentions both a code column and a qty column). This
    # protects against files that have a title row above the real headers.
    header_row_idx = None
    headers = []
    max_scan = min(10, ws.max_row or 1)
    for r in range(1, max_scan + 1):
        row_vals = [clean(c.value) for c in ws[r]]
        joined = ' '.join(row_vals)
        if 'code' in joined and ('qty' in joined or 'quantity' in joined):
            header_row_idx = r
            headers = row_vals
            break
    if header_row_idx is None:
        header_row_idx = 1
        headers = [clean(c.value) for c in ws[1]]

    def find_col(all_of=None, any_of=None):
        for i, h in enumerate(headers):
            if not h:
                continue
            if all_of and all(k in h for k in all_of):
                return i
            if any_of and any(k in h for k in any_of):
                return i
        return None

    code_col = find_col(all_of=['item', 'code']) or find_col(any_of=['code'])
    qty_col = find_col(any_of=['qty', 'quantity'])
    mpn_col = find_col(any_of=['mpn', 'part no', 'partno', 'part number'])
    manu_col = find_col(any_of=['manufact', 'make', 'brand'])
    desc_col = find_col(any_of=['desc', 'description'])

    rows = []
    for r in range(header_row_idx + 1, (ws.max_row or header_row_idx) + 1):
        row_cells = ws[r]

        def val(col_idx):
            if col_idx is None or col_idx >= len(row_cells):
                return None
            return row_cells[col_idx].value

        item_code = val(code_col)
        qty = val(qty_col)
        if item_code is None or qty is None or str(item_code).strip() == '':
            continue
        try:
            qty = float(qty)
        except (TypeError, ValueError):
            continue
        if qty <= 0:
            continue

        mpn_val = val(mpn_col)
        manu_val = val(manu_col)
        desc_val = val(desc_col)

        rows.append({
            'item_code': str(item_code).strip(),
            'qty': qty,
            'mpn': str(mpn_val).strip() if mpn_val not in (None, '') else None,
            'manufacturer_name': str(manu_val).strip() if manu_val not in (None, '') else None,
            'desc': str(desc_val).strip() if desc_val not in (None, '') else None,
        })
    return rows, sheet_name


def calculate_requirements_from_bom(cursor, wo_id, bom_id, qty_to_build):
    """Refreshes ONLY the BOM-sourced rows in work_order_item for this WO.

    Rows are tagged with bom_item_id when they come straight from the BOM.
    Rows added via purchase-sheet import or the manual 'Add new item' form
    always have bom_item_id = NULL, and this function never deletes or
    touches those — only rows tied to a bom_item_id are ever removed or
    rewritten here, so recalculating never wipes out imported/manual items."""
    # ADDED: Check if requirements table is already populated.
    # If it is, only recalculate the qty_required of the existing items and do not pull in new BOM items.
    cursor.execute("SELECT COUNT(*) as cnt FROM work_order_item WHERE wo_id = %s", (wo_id,))
    if cursor.fetchone()['cnt'] > 0:
        cursor.execute("SELECT wo_item_id, bom_qty FROM work_order_item WHERE wo_id = %s", (wo_id,))
        existing_items = cursor.fetchall()
        for row in existing_items:
            bom_qty = float(row['bom_qty'] or 0)
            qty_required = round(bom_qty * float(qty_to_build), 3)
            cursor.execute("UPDATE work_order_item SET qty_required = %s WHERE wo_item_id = %s",
                           (qty_required, row['wo_item_id']))
        return len(existing_items)

    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    cursor.execute("""SELECT bi.bom_item_id, bi.item_id, bi.qty, COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name
                     FROM bom_item bi
                     JOIN items i ON bi.item_id = i.item_id
                     LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                     WHERE bi.bom_id = %s AND bi.is_deleted = 0""", (bom_id,))
    bom_items = cursor.fetchall()
    current_bom_item_ids = [bi['bom_item_id'] for bi in bom_items]

    # Remove stale BOM-linked rows only (items that used to be in this BOM
    # but have since been removed from it). Anything with bom_item_id IS
    # NULL (purchase-sheet imports, manual adds) is never touched.
    if current_bom_item_ids:
        placeholders = ', '.join(['%s'] * len(current_bom_item_ids))
        cursor.execute(f"""DELETE FROM work_order_item
                          WHERE wo_id = %s AND bom_item_id IS NOT NULL
                          AND bom_item_id NOT IN ({placeholders})""",
                       (wo_id, *current_bom_item_ids))
    else:
        cursor.execute("""DELETE FROM work_order_item
                         WHERE wo_id = %s AND bom_item_id IS NOT NULL""", (wo_id,))

    for bi in bom_items:
        qty_required = round(float(bi['qty']) * float(qty_to_build), 3)
        # cursor.execute("""INSERT INTO work_order_item
        #                  (wo_id, item_id, bom_item_id, bom_qty, qty_required, qty_issued, qty_returned, manufacturer_name)
        #                  VALUES (%s, %s, %s, %s, %s, 0, 0, %s)
        #                  ON DUPLICATE KEY UPDATE
        #                      bom_item_id = VALUES(bom_item_id),
        #                      bom_qty = VALUES(bom_qty),
        #                      qty_required = VALUES(qty_required),
        #                      manufacturer_name = COALESCE(VALUES(manufacturer_name), manufacturer_name)""",
        #                (wo_id, bi['item_id'], bi['bom_item_id'], bi['qty'], qty_required, bi.get('manufacturer_name')))
        #
        # added Check if item already exists to prevent duplicates
        cursor.execute("SELECT wo_item_id FROM work_order_item WHERE wo_id = %s AND item_id = %s", (wo_id, bi['item_id']))
        existing = cursor.fetchone()
        if existing:
            cursor.execute("""UPDATE work_order_item
                              SET bom_item_id = %s, bom_qty = %s, qty_required = %s,
                                  manufacturer_name = COALESCE(%s, manufacturer_name)
                              WHERE wo_item_id = %s""",
                           (bi['bom_item_id'], bi['qty'], qty_required, bi.get('manufacturer_name'), existing['wo_item_id']))
        else:
            cursor.execute("""INSERT INTO work_order_item
                             (wo_id, item_id, bom_item_id, bom_qty, qty_required, qty_issued, qty_returned, manufacturer_name)
                             VALUES (%s, %s, %s, %s, %s, 0, 0, %s)""",
                           (wo_id, bi['item_id'], bi['bom_item_id'], bi['qty'], qty_required, bi.get('manufacturer_name')))
    return len(bom_items)


def get_table_columns(cursor, table_name):
    """Introspects live column metadata for a table in the current database.
    Used so we never hard-code the 'items' schema and never risk an
    ALTER-TABLE-style change — we only ever read structure and INSERT/UPDATE
    rows within columns that already exist."""
    cursor.execute("""
        SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, COLUMN_DEFAULT, EXTRA
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = %s
        ORDER BY ORDINAL_POSITION
    """, (table_name,))
    return cursor.fetchall()


_NUMERIC_TYPES = {'int', 'smallint', 'tinyint', 'mediumint', 'bigint',
                   'decimal', 'float', 'double'}


def _safe_default_for_column(col):
    """Picks a harmless value for a NOT NULL column we have no real data
    for, so auto-created items never blow up on a missing-required-field
    error. Returns None if the column should simply be left out of the
    INSERT (nullable, has its own default, or is auto-increment)."""
    if col['IS_NULLABLE'] == 'YES':
        return None
    if col['COLUMN_DEFAULT'] is not None:
        return None
    if 'auto_increment' in (col['EXTRA'] or ''):
        return None
    dtype = (col['DATA_TYPE'] or '').lower()
    if dtype in _NUMERIC_TYPES:
        return 0
    if dtype == 'date':
        return date.today()
    if dtype in ('timestamp', 'datetime'):
        return None  # leave out; if truly NOT NULL with no default this is rare
    return ''  # varchar / text / etc.


def ensure_item_exists(cursor, item_code, desc_val=None, mpn_val=None, manufacturer_val=None): # remove
    """Looks up an item by item_code.

    - If it already exists (even soft-deleted), it is reused and
      reactivated if needed. Any master fields that are currently BLANK
      (desc / mpn / manufacturer_name) get auto-filled from the purchase
      sheet — existing data is never overwritten.
    - If it truly doesn't exist yet, it is auto-created in `items` using
      whatever the purchase sheet gave us, so NOTHING is ever skipped.

    Returns (item_id, was_created: bool).
    """
    item_code = (item_code or '').strip()
    if item_code.startswith("'"):
        item_code = item_code[1:]
    cursor.execute("SELECT * FROM items WHERE item_code = %s", (item_code,))
    existing = cursor.fetchone()

    mfr_id_to_use = None
    if manufacturer_val and manufacturer_val.strip():
        m_name = manufacturer_val.strip()
        cursor.execute("SELECT mfr_id FROM manufacturer WHERE LOWER(mfr_short_name) = %s OR LOWER(mfr_full_name) = %s LIMIT 1", (m_name.lower(), m_name.lower()))
        mfr_row = cursor.fetchone()
        if mfr_row:
            mfr_id_to_use = mfr_row['mfr_id']
        else:
            cursor.execute("INSERT INTO manufacturer (mfr_short_name, mfr_full_name) VALUES (%s, %s)", (m_name, m_name))
            mfr_id_to_use = cursor.lastrowid

    if existing:
        if existing.get('is_deleted'):
            cursor.execute("UPDATE items SET is_deleted = 0 WHERE item_id = %s",
                           (existing['item_id'],))

        fill_updates = {}
        if 'desc' in existing and not (existing.get('desc') or '').strip() and desc_val:
            fill_updates['desc'] = desc_val[:100]
        if 'mpn' in existing and not (existing.get('mpn') or '').strip() and mpn_val:
            fill_updates['mpn'] = mpn_val[:100]
        if not existing.get('mfr_id') and mfr_id_to_use:
            fill_updates['mfr_id'] = mfr_id_to_use
        if fill_updates:
            set_sql = ', '.join(f"`{k}` = %s" for k in fill_updates)
            cursor.execute(f"UPDATE items SET {set_sql} WHERE item_id = %s",
                           list(fill_updates.values()) + [existing['item_id']])
        return existing['item_id'], False

    # Not in the master at all — auto-create it instead of skipping the row.
    columns = get_table_columns(cursor, 'items')
    col_names = {c['COLUMN_NAME'] for c in columns}

    values = {'item_code': item_code}
    if 'desc' in col_names:
        values['desc'] = (desc_val or item_code)[:100]
    if 'mpn' in col_names and mpn_val:
        values['mpn'] = mpn_val[:100]
    if 'mfr_id' in col_names and mfr_id_to_use:
        values['mfr_id'] = mfr_id_to_use
    if 'is_deleted' in col_names:
        values['is_deleted'] = 0

    for col in columns:
        name = col['COLUMN_NAME']
        if name in values or name == 'item_id':
            continue
        default = _safe_default_for_column(col)
        if default is not None:
            values[name] = default

    cols_sql = ', '.join(f"`{k}`" for k in values)
    placeholders = ', '.join(['%s'] * len(values))
    cursor.execute(f"INSERT INTO items ({cols_sql}) VALUES ({placeholders})",
                 list(values.values()))
    return cursor.lastrowid, True


def apply_purchase_sheet_import(cursor, wo_id, qty_to_build, excel_file):
    """Reads the purchase sheet and upserts rows into work_order_item,
    auto-calculating qty_required = bom_qty * qty_to_build.
    NOTHING is ever skipped: if a row's item_code isn't in the `items`
    master yet, it gets auto-created there (using desc/mpn/manufacturer
    straight from the sheet) so it can still be added to the work order.
    Rows that WOULD have been dropped are now imported + flagged as
    newly created, so Erin can review/complete them afterward."""
    # ADDED: Map Excel items to bom_item_id if they exist in the work order's BOM
    cursor.execute("SELECT bom_id FROM work_order WHERE wo_id = %s", (wo_id,))
    wo_row = cursor.fetchone()
    bom_id = wo_row['bom_id'] if wo_row else None

    bom_item_map = {}
    if bom_id:
        cursor.execute("SELECT item_id, bom_item_id FROM bom_item WHERE bom_id = %s AND is_deleted = 0", (bom_id,))
        for b_item in cursor.fetchall():
            bom_item_map[b_item['item_id']] = b_item['bom_item_id']

    rows, sheet_used = parse_purchase_sheet(io.BytesIO(excel_file.read()))
    imported = 0
    created = 0
    created_codes = []
    for row in rows:
        item_id, was_created = ensure_item_exists(
            cursor, row['item_code'], row.get('desc'), row.get('mpn'), row.get('manufacturer_name'))
        if was_created:
            created += 1
            created_codes.append(row['item_code'])

        bom_qty = row['qty']
        qty_required = round(bom_qty * float(qty_to_build), 3)
        manufacturer_name = row.get('manufacturer_name')
        
        # cursor.execute("""INSERT INTO work_order_item
        #                  (wo_id, item_id, bom_qty, qty_required, qty_issued, qty_returned, manufacturer_name)
        #                  VALUES (%s, %s, %s, %s, 0, 0, %s)
        #                  ON DUPLICATE KEY UPDATE
        #                      bom_qty = VALUES(bom_qty),
        #                      qty_required = VALUES(qty_required),
        #                      manufacturer_name = COALESCE(VALUES(manufacturer_name), manufacturer_name)""",
        #                (wo_id, item_id, bom_qty, qty_required, manufacturer_name))
        
        # To Check if item already exists to prevent duplicates
        bom_item_id = bom_item_map.get(item_id)
        cursor.execute("SELECT wo_item_id FROM work_order_item WHERE wo_id = %s AND item_id = %s", (wo_id, item_id))
        existing = cursor.fetchone()
        if existing:
            cursor.execute("""UPDATE work_order_item
                              SET bom_qty = %s, qty_required = %s, is_removed = 0,
                                  bom_item_id = COALESCE(%s, bom_item_id),
                                  manufacturer_name = COALESCE(%s, manufacturer_name)
                              WHERE wo_item_id = %s""",
                           (bom_qty, qty_required, bom_item_id, manufacturer_name, existing['wo_item_id']))
        else:
            cursor.execute("""INSERT INTO work_order_item
                             (wo_id, item_id, bom_item_id, bom_qty, qty_required, qty_issued, qty_returned, manufacturer_name)
                             VALUES (%s, %s, %s, %s, %s, 0, 0, %s)""",
                           (wo_id, item_id, bom_item_id, bom_qty, qty_required, manufacturer_name))
        imported += 1
    return imported, created, created_codes, sheet_used


def get_item_stock(cursor, item_id):
    cursor.execute("""
        SELECT COALESCE(SUM(s.physical_availability), 0) as total_in
        FROM storage s WHERE s.item_id = %s
    """, (item_id,))
    total_in = float(cursor.fetchone()['total_in'] or 0)

    try:
        cursor.execute("""
            SELECT COALESCE(SUM(qty_issued), 0) as total_out
            FROM issue_transaction WHERE item_id = %s
        """, (item_id,))
        total_out = float(cursor.fetchone()['total_out'] or 0)
    except:
        total_out = 0.0

    try:
        cursor.execute("""
            SELECT COALESCE(SUM(qty_blocked), 0) as total_blocked
            FROM stock_blocking WHERE item_id = %s AND block_status = 0
        """, (item_id,))
        total_blocked = float(cursor.fetchone()['total_blocked'] or 0)
    except:
        total_blocked = 0.0

    # Currently available unblocked stock in storage = physical store total minus active blocks
    available = max(0.0, round(total_in - total_blocked, 3))
    return {
        'total_in': round(total_in, 3),
        'total_out': round(total_out, 3),
        'blocked': round(total_blocked, 3),
        'available': round(available, 3)
    }

def get_vendor_items(cursor):
    cursor.execute("""
        SELECT iv.vendor_id, iv.item_id, i.item_code, i.`desc`, iv.unit_id
        FROM item_vendor iv
        JOIN items i ON iv.item_id = i.item_id
        WHERE iv.is_deleted = 0 AND i.is_deleted = 0
        ORDER BY i.item_code
    """)
    vendor_items_list = cursor.fetchall()
    
    vendor_items = {}
    for row in vendor_items_list:
        v_id = row['vendor_id']
        if v_id not in vendor_items:
            vendor_items[v_id] = []
        vendor_items[v_id].append({
            'item_id': row['item_id'],
            'item_code': row['item_code'],
            'desc': row['desc'],
            'unit_id': row['unit_id']
        })
    return vendor_items

# ─── PO NUMBER GENERATOR ────────────────────────────────
def generate_po_number(cursor):
    # Find start and end dates of the current financial year
    today = date.today()
    if today.month >= 4:
        fy_start = f"{today.year}-04-01"
        fy_end = f"{today.year + 1}-03-31"
    else:
        fy_start = f"{today.year - 1}-04-01"
        fy_end = f"{today.year}-03-31"
        
    cursor.execute("""
        SELECT MAX(po_number) as max_po 
        FROM purchase_order 
        WHERE date_raised >= %s AND date_raised <= %s
    """, (fy_start, fy_end))
    row = cursor.fetchone()
    if row and row['max_po'] is not None:
        return row['max_po'] + 1
    return 1

# ─── DASHBOARD ──────────────────────────────────────────
# ─── DASHBOARD ──────────────────────────────────────────
@app.route('/')
def dashboard():
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT COUNT(*) as cnt FROM work_order WHERE status = 0")
    wo_open = cursor.fetchone()['cnt']
    cursor.execute("SELECT COUNT(*) as cnt FROM work_order WHERE status = 1")
    wo_inprogress = cursor.fetchone()['cnt']
    cursor.execute("SELECT COUNT(*) as cnt FROM grn")
    grn_total = cursor.fetchone()['cnt']
    cursor.execute("SELECT COUNT(*) as cnt FROM purchase_order")
    po_total = cursor.fetchone()['cnt']
    cursor.execute("SELECT COUNT(*) as cnt FROM purchase_order WHERE po_status IN ('Open', 'Draft')")
    po_open = cursor.fetchone()['cnt']
    cursor.execute("SELECT COUNT(*) as cnt FROM grn_item WHERE qc_status = 'Pending'")
    grn_qc_pending = cursor.fetchone()['cnt']

    # 1. System Alerts
    system_alerts = []

    # Financial Year Alert
    current_fy_label = get_financial_year(cursor=cursor)
    fy_start_month = get_setting('fy_start_month', 'March', cursor=cursor)
    fy_start_year = get_setting('fy_start_year', '2027', cursor=cursor)
    fy_end_month = get_setting('fy_end_month', 'April', cursor=cursor)
    fy_end_year = get_setting('fy_end_year', '2028', cursor=cursor)

    fy_alert_text = f"Current Financial Year: <strong>{fy_start_month} {fy_start_year}</strong> to <strong>{fy_end_month} {fy_end_year}</strong> (Active FY: <strong>{current_fy_label}</strong>)"
    system_alerts.append({
        'type': 'info',
        'dot_class': 'dot-active',
        'title': 'Financial Year Period',
        'text': fy_alert_text,
        'sub': 'All generated WO, PO, and GRN numbers adapt to this financial year.'
    })

    # Policy Expiration Alert
    policy_text = get_setting('policy_line_text', 'COVERED UNDER NEW INDIA ASSURANCE POLICY NO. 67020021200200000030', cursor=cursor)
    policy_exp_str = get_setting('policy_expiration_date', '2026-12-31', cursor=cursor)

    try:
        exp_date = date.fromisoformat(policy_exp_str)
        formatted_exp_date = exp_date.strftime('%d %B %Y')
        is_expired = (date.today() >= exp_date)
    except Exception:
        formatted_exp_date = policy_exp_str
        is_expired = False

    if is_expired:
        policy_alert_text = f"Insurance Policy Line <strong>EXPIRED on {formatted_exp_date}</strong>. Please update policy settings immediately."
        dot_cls = 'dot-danger'
        alert_type = 'danger'
    else:
        policy_alert_text = f"Insurance Policy Line is active and valid until <strong>{formatted_exp_date}</strong>."
        dot_cls = 'dot-success'
        alert_type = 'success'

    system_alerts.append({
        'type': alert_type,
        'dot_class': dot_cls,
        'title': 'Insurance Policy Status',
        'text': policy_alert_text,
        'sub': f'Current Policy: "{policy_text}"'
    })

    # 2. Recent Activities (Logins, WOs, GRNs)
    recent_activities = []

    # Recent user logins
    cursor.execute("SELECT username, full_name, role, last_login FROM users WHERE last_login IS NOT NULL ORDER BY last_login DESC LIMIT 5")
    for u in cursor.fetchall():
        user_display = u['full_name'] or u['username']
        recent_activities.append({
            'icon': '👤',
            'text': f"User <strong>{user_display}</strong> ({u['username']}) logged in",
            'time': u['last_login'].strftime('%d-%m-%Y %I:%M %p') if u['last_login'] else '',
            'dt': u['last_login']
        })

    # Recent WOs created
    cursor.execute("SELECT wo_number, created_at, created_by FROM work_order ORDER BY wo_id DESC LIMIT 3")
    for wo in cursor.fetchall():
        if wo.get('created_at'):
            by_str = f" by {wo['created_by']}" if wo.get('created_by') else ""
            recent_activities.append({
                'icon': '📋',
                'text': f"Work Order <strong>#{wo['wo_number']}</strong> created{by_str}",
                'time': wo['created_at'].strftime('%d-%m-%Y %I:%M %p'),
                'dt': wo['created_at']
            })

    # Recent GRNs created
    cursor.execute("SELECT grn_no, created_at, received_by FROM grn ORDER BY grn_id DESC LIMIT 3")
    for grn in cursor.fetchall():
        if grn.get('created_at'):
            by_str = f" by {grn['received_by']}" if grn.get('received_by') else ""
            recent_activities.append({
                'icon': '📦',
                'text': f"GRN <strong>{grn['grn_no']}</strong> recorded{by_str}",
                'time': grn['created_at'].strftime('%d-%m-%Y %I:%M %p'),
                'dt': grn['created_at']
            })

    # Sort combined activities by datetime descending
    recent_activities.sort(key=lambda x: x['dt'], reverse=True)
    recent_activities = recent_activities[:8]

    cursor.close()
    db.close()
    return render_template('dashboard.html',
        wo_open=wo_open,
        wo_inprogress=wo_inprogress,
        grn_total=grn_total,
        grn_qc_pending=grn_qc_pending,
        po_total=po_total,
        po_open=po_open,
        system_alerts=system_alerts,
        recent_activities=recent_activities
    )

# ─── PURCHASE ORDER LIST ────────────────────────────────
@app.route('/purchase-order')
def purchase_order_list():
    search = request.args.get('search', '')
    status_filter = request.args.get('status', '')
    db = get_db()
    cursor = db.cursor(dictionary=True)

    query = """SELECT po.*, v.short_name as vendor_name, v.full_name as vendor_full_name
               FROM purchase_order po
               JOIN vendor v ON po.vendor_id = v.vendor_id
               WHERE 1=1"""
    params = []

    if search:
        query += " AND (po.po_number LIKE %s OR po.po_status LIKE %s OR v.short_name LIKE %s OR v.full_name LIKE %s)"
        params += [f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%']
    if status_filter != '':
        query += " AND po.po_status = %s"
        params.append(status_filter)

    query += " ORDER BY po.po_id DESC"
    cursor.execute(query, params)
    purchase_orders = cursor.fetchall()

    # Pass formatted PO numbers
    for po in purchase_orders:
        po['formatted_po_number'] = format_po_number(po['po_number'], po['date_raised'], po['po_version_number'])

    # Batch-fetch PO items
    po_ids = [po['po_id'] for po in purchase_orders]
    po_items_map = {}
    if po_ids:
        format_strings = ','.join(['%s'] * len(po_ids))
        # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
        cursor.execute(f"""
            SELECT pi.*, i.item_code, i.desc as item_desc, i.mpn,
                   COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name, u.unit_short_name as unit,
                   h.hsn_code, h.tax_rate, h.cgst, h.sgst, h.igst
            FROM po_items pi
            JOIN items i ON pi.item_id = i.item_id
            LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
            LEFT JOIN units u ON pi.unit_id = u.unit_id
            LEFT JOIN hsn h ON pi.hsn_id = h.hsn_id
            WHERE pi.po_id IN ({format_strings})
        """, tuple(po_ids))
        all_items = cursor.fetchall()
        for item in all_items:
            po_id = item['po_id']
            if po_id not in po_items_map:
                po_items_map[po_id] = []
            po_items_map[po_id].append(item)

    cursor.close()
    db.close()
    return render_template('purchase_order_list.html',
                           purchase_orders=purchase_orders,
                           search=search,
                           status_filter=status_filter,
                           po_items_map=po_items_map)

# ─── PURCHASE ORDER ADD ─────────────────────────────────
@app.route('/purchase-order/add', methods=['GET', 'POST'])
def purchase_order_add():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        po_number = request.form['po_number']
        vendor_input = (request.form.get('vendor_id') or '').strip()
        vendor_id = None

        if vendor_input:
            if vendor_input.isdigit():
                cursor.execute("SELECT vendor_id FROM vendor WHERE vendor_id = %s AND is_deleted = 0", (int(vendor_input),))
                v_row = cursor.fetchone()
                if v_row:
                    vendor_id = v_row['vendor_id']
            
            if not vendor_id:
                # Search by short_name or full_name
                cursor.execute("SELECT vendor_id FROM vendor WHERE (short_name = %s OR full_name = %s) AND is_deleted = 0 LIMIT 1", (vendor_input, vendor_input))
                v_row = cursor.fetchone()
                if v_row:
                    vendor_id = v_row['vendor_id']
                else:
                    # Auto-create new vendor entry
                    import time
                    gst_dummy = f"URP-{int(time.time())}"
                    short_name = vendor_input[:50]
                    full_name = vendor_input[:255]
                    cursor.execute("""
                        INSERT INTO vendor (short_name, full_name, gst_no, is_deleted)
                        VALUES (%s, %s, %s, 0)
                    """, (short_name, full_name, gst_dummy))
                    vendor_id = cursor.lastrowid

        date_raised = request.form.get('date_raised') or None
        remarks = request.form.get('remarks') or None
        currency = request.form.get('currency', 'INR').strip()
        custom_terms = request.form.get('custom_terms') or None
        payment_terms = request.form.get('payment_terms') or None
        
        # New PO starts as version 0 and Draft status
        po_version_number = '0'
        po_status = 'Draft'
        raised_by = session.get('username') or 'admin'

        cursor.execute("""INSERT INTO purchase_order 
                         (po_number, po_version_number, vendor_id, date_raised, raised_by,
                          po_status, remarks, currency, custom_terms, payment_terms)
                         VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                       (po_number, po_version_number, vendor_id, date_raised, raised_by,
                        po_status, remarks, currency, custom_terms, payment_terms))
        po_id = cursor.lastrowid
        
        # Save lines to po_items
        item_ids = request.form.getlist('item_id[]')
        qty_ordered_list = request.form.getlist('qty_ordered[]')
        unit_price_list = request.form.getlist('unit_price[]')
        unit_id_list = request.form.getlist('unit_id[]')
        
        # Default HSN (excluding 999999)
        cursor.execute("SELECT hsn_id FROM hsn WHERE hsn_code != '999999' AND is_deleted = 0 ORDER BY hsn_id ASC LIMIT 1")
        hsn_row = cursor.fetchone()
        hsn_id = hsn_row['hsn_id'] if hsn_row else 1

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty = float(qty_ordered_list[i] or 0)
            price = float(unit_price_list[i] or 0)
            unit_id = int(unit_id_list[i])
            
            # Migrated from items.hsn_id — now sourced via hsn_items junction table
            cursor.execute("SELECT hsn_id FROM hsn_items WHERE item_id = %s", (item_id,))
            item_row = cursor.fetchone()
            line_hsn_id = item_row['hsn_id'] if (item_row and item_row.get('hsn_id')) else hsn_id

            cursor.execute("""INSERT INTO po_items 
                             (po_id, item_id, hsn_id, qty_ordered, unit_price, unit_id, status)
                             VALUES (%s, %s, %s, %s, %s, %s, 0)""",
                           (po_id, item_id, line_hsn_id, qty, price, unit_id))

        db.commit()
        cursor.close()
        db.close()
        
        formatted_po = format_po_number(int(po_number), date_raised, po_version_number)
        flash(f'Purchase Order {formatted_po} created successfully.', 'success')
        return redirect(url_for('purchase_order_list'))

    # Suggest next sequence number in current financial year
    next_po = generate_po_number(cursor)
    next_formatted_po = format_po_number(next_po, date.today(), 0)

    # Get vendors
    cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted=0 ORDER BY short_name")
    vendors = cursor.fetchall()
    
    # Get items for dropdown
    cursor.execute("SELECT item_id, item_code, `desc`, unit_id FROM items WHERE is_deleted=0 ORDER BY item_code")
    items = cursor.fetchall()
    
    # Get units for dropdown
    cursor.execute("SELECT unit_id, unit_short_name FROM units ORDER BY unit_name")
    units = cursor.fetchall()

    vendor_items = get_vendor_items(cursor)

    cursor.close()
    db.close()
    return render_template('purchase_order_add.html', 
                           vendors=vendors, 
                           next_po=next_po, 
                           next_formatted_po=next_formatted_po, 
                           today=date.today(),
                           items=items,
                           units=units,
                           vendor_items=vendor_items)

# ─── GET VENDOR PRICING FOR AUTOFILL ────────────────────
@app.route('/purchase-order/get-vendor-pricing')
def get_vendor_pricing():
    vendor_id = request.args.get('vendor_id')
    item_id = request.args.get('item_id')
    if not vendor_id or not item_id:
        return jsonify(success=False, error="Missing parameters")
    
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    cursor.execute("""
        SELECT iv.unit_id, ivp.qty, ivp.price
        FROM item_vendor iv
        LEFT JOIN item_vendor_pricing ivp ON iv.item_vendor_id = ivp.item_vendor_id AND ivp.is_deleted = 0
        WHERE iv.item_id = %s AND iv.vendor_id = %s AND iv.is_deleted = 0
        ORDER BY ivp.qty ASC
        LIMIT 1
    """, (item_id, vendor_id))
    row = cursor.fetchone()
    
    cursor.close()
    db.close()
    
    if row:
        return jsonify(
            success=True,
            unit_id=row['unit_id'],
            qty=float(row['qty']) if row['qty'] is not None else None,
            price=float(row['price']) if row['price'] is not None else None
        )
    return jsonify(success=False, error="No pricing found")


# ─── PURCHASE ORDER DETAIL ──────────────────────────────
@app.route('/purchase-order/<int:po_id>')
def purchase_order_detail(po_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("""SELECT po.*, v.short_name as vendor_name, v.full_name as vendor_full_name,
                             v.gst_no, v.ph_no, v.address_line_1, v.city, v.pincode, v.country
                      FROM purchase_order po
                      JOIN vendor v ON po.vendor_id = v.vendor_id
                      WHERE po.po_id = %s""", (po_id,))
    po = cursor.fetchone()
    if not po:
        cursor.close()
        db.close()
        return "Purchase Order not found", 404

    # Format PO number
    po['formatted_po_number'] = format_po_number(po['po_number'], po['date_raised'], po['po_version_number'])

    # Fetch PO items
    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    cursor.execute("""SELECT pi.*, i.item_code, i.desc as item_desc, i.mpn,
                             COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name, u.unit_short_name as unit,
                             h.hsn_code, h.tax_rate, h.cgst, h.sgst, h.igst
                      FROM po_items pi
                      JOIN items i ON pi.item_id = i.item_id
                      LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                      LEFT JOIN units u ON pi.unit_id = u.unit_id
                      LEFT JOIN hsn h ON pi.hsn_id = h.hsn_id
                      WHERE pi.po_id = %s""", (po_id,))
    po_items = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('purchase_order_detail.html', po=po, po_items=po_items)

# ─── PURCHASE ORDER ACTION: GENERATE PO PDF ─────────────
@app.route('/purchase-order/<int:po_id>/po-pdf')
def purchase_order_po_pdf(po_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("""SELECT po.*, v.short_name, v.full_name, v.gst_no, v.ph_no,
                             v.address_line_1, v.city, v.pincode, v.country
                      FROM purchase_order po
                      JOIN vendor v ON po.vendor_id = v.vendor_id
                      WHERE po.po_id = %s""", (po_id,))
    po = cursor.fetchone()
    if not po:
        cursor.close()
        db.close()
        return "Purchase Order not found", 404

    po['formatted_po_number'] = format_po_number(po['po_number'], po['date_raised'], po['po_version_number'])

    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    cursor.execute("""SELECT pi.*, i.item_code, i.desc as item_desc, i.mpn,
                             COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name, u.unit_short_name as unit,
                             h.hsn_code, h.tax_rate, h.cgst, h.sgst, h.igst
                      FROM po_items pi
                      JOIN items i ON pi.item_id = i.item_id
                      LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                      LEFT JOIN units u ON pi.unit_id = u.unit_id
                      LEFT JOIN hsn h ON pi.hsn_id = h.hsn_id
                      WHERE pi.po_id = %s""", (po_id,))
    po_items = cursor.fetchall()

    cursor.close()
    db.close()

    custom_terms = request.args.get('terms')
    if custom_terms is None:
        custom_terms = po.get('custom_terms')

    payment_terms = request.args.get('payment_terms')
    if payment_terms is None:
        payment_terms = po.get('payment_terms')

    mode = request.args.get('mode', '')
    hide_signature = (mode == 'review')

    # Hide signature if PO is not Approved and they are not in the approve preview mode
    if po.get('po_status') != 'Approved' and mode != 'approve':
        hide_signature = True

    # Override approved_by and approved_date dynamically for approval preview if not approved in DB yet
    if mode == 'approve' and not po.get('approved_by'):
        from datetime import date
        po['approved_by'] = session.get('full_name') or session.get('username') or 'admin'
        po['approved_date'] = date.today()

    policy_line_text = get_setting('policy_line_text', 'COVERED UNDER NEW INDIA ASSURANCE POLICY NO. 67020021200200000030', cursor=cursor)

    from po_invoice_pdf import build_po_pdf
    pdf_buffer = build_po_pdf(po, po_items, po, custom_terms=custom_terms, payment_terms=payment_terms, hide_signature=hide_signature, policy_line=policy_line_text)

    clean_num = str(po['formatted_po_number']).replace('/', '_')
    filename = f"PO_{clean_num}.pdf"
    return send_file(
        pdf_buffer,
        mimetype='application/pdf',
        as_attachment=False,
        download_name=filename
    )

# ─── PURCHASE ORDER EDIT ────────────────────────────────
@app.route('/purchase-order/<int:po_id>/edit', methods=['GET', 'POST'])
def purchase_order_edit(po_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("SELECT * FROM purchase_order WHERE po_id = %s", (po_id,))
    po = cursor.fetchone()
    if not po:
        cursor.close()
        db.close()
        return "Purchase Order not found", 404

    # Lock Guard
    if po['is_locked'] == 1:
        cursor.close()
        db.close()
        flash("This Purchase Order is locked. Unlock it to edit/amend.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    # We do NOT update the DB status to 'Draft' on GET, we just override it locally 
    # so the edit page selects 'Draft'. The actual database status remains 'Approved' 
    # until they save the form (POST).
    is_previously_approved = (po['po_status'] == 'Approved')

    if request.method == 'POST':
        vendor_id = request.form['vendor_id']
        date_raised = request.form.get('date_raised') or None
        remarks = request.form.get('remarks') or None
        po_status = request.form.get('po_status') or 'Draft'
        
        reviewed_date = request.form.get('reviewed_date') or None
        reviewed_by = request.form.get('reviewed_by') or None
        approved_date = request.form.get('approved_date') or None
        approved_by = request.form.get('approved_by') or None
        cancelled_by = request.form.get('cancelled_by') or None
        cancel_type = request.form.get('cancel_type') or None
        cancelled_at = request.form.get('cancelled_at') or None

        # Determine version update:
        # Increment version only if previous status was 'Approved'
        po_version_number = po['po_version_number'] or '0'
        if is_previously_approved:
            try:
                po_version_number = str(int(po_version_number) + 1)
            except ValueError:
                po_version_number = '1'

        # Automatic cancelled_at timestamp if status is Cancelled
        if po_status == 'Cancelled' and not cancelled_at:
            from datetime import datetime
            cancelled_at = datetime.now()
        elif po_status != 'Cancelled':
            cancelled_at = None
            cancelled_by = None
            cancel_type = None

        currency = request.form.get('currency', 'INR').strip()
        custom_terms = request.form.get('custom_terms') or None
        payment_terms = request.form.get('payment_terms') or None

        cursor.execute("""UPDATE purchase_order SET
                            vendor_id = %s, date_raised = %s, po_status = %s,
                            reviewed_date = %s, reviewed_by = %s, approved_date = %s, approved_by = %s,
                            cancelled_by = %s, cancel_type = %s, cancelled_at = %s, remarks = %s,
                            po_version_number = %s, currency = %s, custom_terms = %s, payment_terms = %s
                          WHERE po_id = %s""",
                       (vendor_id, date_raised, po_status,
                        reviewed_date, reviewed_by, approved_date, approved_by,
                        cancelled_by, cancel_type, cancelled_at, remarks,
                        po_version_number, currency, custom_terms, payment_terms, po_id))
        
        # Clear existing items
        cursor.execute("DELETE FROM po_items WHERE po_id = %s", (po_id,))
        
        # Save lines to po_items
        item_ids = request.form.getlist('item_id[]')
        qty_ordered_list = request.form.getlist('qty_ordered[]')
        unit_price_list = request.form.getlist('unit_price[]')
        unit_id_list = request.form.getlist('unit_id[]')
        
        # Default HSN (excluding 999999)
        cursor.execute("SELECT hsn_id FROM hsn WHERE hsn_code != '999999' AND is_deleted = 0 ORDER BY hsn_id ASC LIMIT 1")
        hsn_row = cursor.fetchone()
        hsn_id = hsn_row['hsn_id'] if hsn_row else 1

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty = float(qty_ordered_list[i] or 0)
            price = float(unit_price_list[i] or 0)
            unit_id = int(unit_id_list[i])
            
            # Migrated from items.hsn_id — now sourced via hsn_items junction table
            cursor.execute("SELECT hsn_id FROM hsn_items WHERE item_id = %s", (item_id,))
            item_row = cursor.fetchone()
            line_hsn_id = item_row['hsn_id'] if (item_row and item_row.get('hsn_id')) else hsn_id

            cursor.execute("""INSERT INTO po_items 
                             (po_id, item_id, hsn_id, qty_ordered, unit_price, unit_id, status)
                             VALUES (%s, %s, %s, %s, %s, %s, 0)""",
                           (po_id, item_id, line_hsn_id, qty, price, unit_id))

        db.commit()
        cursor.close()
        db.close()
        
        formatted_po = format_po_number(po['po_number'], date_raised, po_version_number)
        flash(f'Purchase Order {formatted_po} updated successfully.', 'success')
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    # GET route: query lists
    cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted=0 ORDER BY short_name")
    vendors = cursor.fetchall()
    
    cursor.execute("SELECT item_id, item_code, `desc`, unit_id FROM items WHERE is_deleted=0 ORDER BY item_code")
    items = cursor.fetchall()
    
    cursor.execute("SELECT unit_id, unit_short_name FROM units ORDER BY unit_name")
    units = cursor.fetchall()
    
    cursor.execute("SELECT * FROM po_items WHERE po_id = %s", (po_id,))
    po_items = cursor.fetchall()

    po['formatted_po_number'] = format_po_number(po['po_number'], po['date_raised'], po['po_version_number'])

    ver_val = 0
    if po['po_version_number']:
        try:
            ver_val = int(float(po['po_version_number']))
        except ValueError:
            pass
            
    is_amend = is_previously_approved or (ver_val > 0)

    vendor_items = get_vendor_items(cursor)

    if is_previously_approved:
        po['po_status'] = 'Draft'

    cursor.close()
    db.close()
    return render_template('purchase_order_edit.html', po=po, vendors=vendors, items=items, units=units, po_items=po_items, is_amend=is_amend, vendor_items=vendor_items)


# ─── PURCHASE ORDER ACTION: LOCK ────────────────────────
@app.route('/purchase-order/<int:po_id>/lock', methods=['POST'])
def purchase_order_lock(po_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute("UPDATE purchase_order SET is_locked = 1 WHERE po_id = %s", (po_id,))
    db.commit()
    cursor.close()
    db.close()
    flash("Purchase Order locked. Edit access suspended.", "success")
    return redirect(request.referrer or url_for('purchase_order_detail', po_id=po_id))


# ─── PURCHASE ORDER ACTION: UNLOCK ──────────────────────
@app.route('/purchase-order/<int:po_id>/unlock', methods=['POST'])
def purchase_order_unlock(po_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute("UPDATE purchase_order SET is_locked = 0 WHERE po_id = %s", (po_id,))
    db.commit()
    cursor.close()
    db.close()
    flash("Purchase Order unlocked. Edit access restored.", "success")
    return redirect(request.referrer or url_for('purchase_order_detail', po_id=po_id))


# ─── PURCHASE ORDER ACTION: REVIEW ──────────────────────
@app.route('/purchase-order/<int:po_id>/review', methods=['POST'])
def purchase_order_review(po_id):
    from datetime import date
    db = get_db()
    cursor = db.cursor()
    reviewer = session.get('full_name') or session.get('username') or 'admin'
    cursor.execute("""
        UPDATE purchase_order 
        SET po_status = 'Reviewed', reviewed_date = %s, reviewed_by = %s
        WHERE po_id = %s
    """, (date.today(), reviewer, po_id))
    db.commit()
    cursor.close()
    db.close()
    flash("Purchase Order marked as Reviewed.", "success")
    return redirect(url_for('purchase_order_detail', po_id=po_id))


# ─── PURCHASE ORDER ACTION: APPROVE ─────────────────────
@app.route('/purchase-order/<int:po_id>/approve', methods=['POST'])
def purchase_order_approve(po_id):
    if session.get('role') not in ['admin', 'approver']:
        flash("You do not have authority to approve this Purchase Order.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))
        
    approver = session.get('full_name') or session.get('username') or 'admin'
    from datetime import date
    db = get_db()
    cursor = db.cursor()
    cursor.execute("""
        UPDATE purchase_order 
        SET po_status = 'Approved', approved_date = %s, approved_by = %s
        WHERE po_id = %s
    """, (date.today(), approver, po_id))

    db.commit()
    cursor.close()
    db.close()
    flash("Purchase Order approved successfully.", "success")
    return redirect(url_for('purchase_order_detail', po_id=po_id))


# ─── WORK ORDER LIST ────────────────────────────────────
@app.route('/work-order')
def work_order_list():
    search = request.args.get('search', '')
    status_filter = request.args.get('status', '')
    pending_only = request.args.get('pending_only', '')
    db = get_db()
    cursor = db.cursor(dictionary=True)

    query = """SELECT wo.*, b.bom_code, b.description as bom_desc, b.bom_id
               FROM work_order wo
               JOIN bom b ON wo.bom_id = b.bom_id
               WHERE 1=1"""
    params = []

    if search:
        query += " AND (wo.wo_number LIKE %s OR wo.description LIKE %s OR b.bom_code LIKE %s)"
        params += [f'%{search}%', f'%{search}%', f'%{search}%']
    if status_filter != '':
        query += " AND wo.status = %s"
        params.append(status_filter)

    query += " ORDER BY wo.wo_id DESC"
    cursor.execute(query, params)
    work_orders = cursor.fetchall()

    for wo in work_orders:
        cursor.execute("""
            SELECT COUNT(*) as total,
                   SUM(CASE WHEN qty_required > qty_issued THEN 1 ELSE 0 END) as pending
            FROM work_order_item WHERE wo_id = %s AND COALESCE(is_removed, 0) = 0
        """, (wo['wo_id'],))
        req = cursor.fetchone()
        wo['req_total'] = req['total'] or 0
        wo['req_pending'] = req['pending'] or 0

    if pending_only:
        work_orders = [w for w in work_orders if w['req_pending'] > 0]

    cursor.close()
    db.close()
    return render_template('work_order_list.html',
                           work_orders=work_orders,
                           search=search,
                           status_filter=status_filter,
                           pending_only=pending_only,
                           status_map=STATUS_MAP)

@app.route('/work-order/parse-bom-file', methods=['POST'])
def parse_bom_file():
    if 'bom_file' not in request.files:
        return jsonify({'status': 'error', 'message': 'No file uploaded'})
    
    file = request.files['bom_file']
    if not file or file.filename == '':
        return jsonify({'status': 'error', 'message': 'No file selected'})
    
    try:
        file_bytes = file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        bom_code, sheet_name = get_bom_from_sheets(wb.sheetnames)
        if not bom_code:
            return jsonify({
                'status': 'error',
                'message': "Could not find a sheet starting with 'Purchase_' or 'Purchase-' to extract the BOM number."
            })
        
        # Cross reference in DB (tribi_db)
        db = get_db()
        cursor = db.cursor(dictionary=True)
        cursor.execute("SELECT bom_id, bom_code, description FROM bom WHERE bom_code = %s AND is_deleted = 0", (bom_code,))
        bom_row = cursor.fetchone()
        cursor.close()
        db.close()
        
        if bom_row:
            return jsonify({
                'status': 'success',
                'bom_id': bom_row['bom_id'],
                'bom_code': bom_row['bom_code'],
                'description': bom_row['description'] or 'No description available',
                'sheet_name': sheet_name
            })
        else:
            return jsonify({
                'status': 'not_found',
                'bom_code': bom_code,
                'sheet_name': sheet_name,
                'message': f"BOM number '{bom_code}' extracted from sheet '{sheet_name}' was not found in database 'tribi_db'."
            })
            
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error reading Excel file: {str(e)}"})


# ─── WORK ORDER ADD ─────────────────────────────────────
@app.route('/work-order/add', methods=['GET', 'POST'])
def work_order_add():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        qty_to_build = request.form['qty_to_build']
        description = request.form.get('description', '')
        created_by = request.form.get('created_by', '')
        remarks = request.form.get('remarks', '')
        action = request.form.get('action', 'draft')

        bom_file = request.files.get('bom_file')
        if not bom_file or bom_file.filename == '':
            flash('Please upload an Excel BOM file.', 'error')
            cursor.close()
            db.close()
            return render_template('work_order_add.html')

        try:
            file_bytes = bom_file.read()
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            bom_code, sheet_name = get_bom_from_sheets(wb.sheetnames)
            if not bom_code:
                flash("Could not find a sheet starting with 'Purchase_' or 'Purchase-' to extract the BOM number from the uploaded file.", 'error')
                cursor.close()
                db.close()
                return render_template('work_order_add.html')

            # Fetch BOM details from tribi_db
            cursor.execute("SELECT bom_id, description FROM bom WHERE bom_code = %s AND is_deleted = 0", (bom_code,))
            bom_row = cursor.fetchone()
            if not bom_row:
                flash(f"BOM number '{bom_code}' extracted from sheet '{sheet_name}' was not found in database 'tribi_db'.", 'error')
                cursor.close()
                db.close()
                return render_template('work_order_add.html')
            
            bom_id = bom_row['bom_id']
            bom_desc = bom_row['description'] or ''
            
            # If description was empty, we can use the BOM's description or just keep the user entered one
            if not description and bom_desc:
                description = bom_desc

        except Exception as e:
            flash(f"Error parsing uploaded Excel file: {str(e)}", 'error')
            cursor.close()
            db.close()
            return render_template('work_order_add.html')

        wo_number = generate_wo_number(cursor)
        cursor.execute("""INSERT INTO work_order 
                         (wo_number, bom_id, qty_to_build, description, created_by, status, remarks)
                         VALUES (%s, %s, %s, %s, %s, 0, %s)""",
                       (wo_number, bom_id, qty_to_build, description, created_by, remarks))
        wo_id = cursor.lastrowid

        if action == 'calculate':
            # Run the purchase sheet import using the uploaded Excel file
            excel_file_stream = io.BytesIO(file_bytes)
            imported, created, created_codes, sheet_used = apply_purchase_sheet_import(
                cursor, wo_id, qty_to_build, excel_file_stream)
            db.commit()
            cursor.close()
            db.close()
            if imported:
                msg = f"Work order created. Imported {imported} item(s) from sheet '{sheet_used}'."
                if created:
                    preview = ', '.join(created_codes[:8])
                    if len(created_codes) > 8:
                        preview += f", +{len(created_codes) - 8} more"
                    msg += (f" {created} item(s) weren't in the item master yet, so they were "
                            f"auto-added ({preview}) — please review their details on the Items page.")
                flash(msg, 'success')
            else:
                flash(f"Work order created from BOM, but sheet '{sheet_used}' had no usable rows "
                      f"(check that it has item code and qty columns).", 'error')
            return redirect(url_for('work_order_detail', wo_id=wo_id))

        db.commit()
        cursor.close()
        db.close()
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    cursor.close()
    db.close()
    return render_template('work_order_add.html')

# ─── WORK ORDER DETAIL ──────────────────────────────────
@app.route('/work-order/<int:wo_id>')
def work_order_detail(wo_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("""SELECT wo.*, b.bom_code, b.description as bom_desc, b.bom_id
                     FROM work_order wo
                     JOIN bom b ON wo.bom_id = b.bom_id
                     WHERE wo.wo_id = %s""", (wo_id,))
    wo = cursor.fetchone()
    if not wo:
        cursor.close()
        db.close()
        return "Work order not found", 404

    # Auto-calculate requirements from BOM on load
    if wo['bom_id']:
        calculate_requirements_from_bom(cursor, wo_id, wo['bom_id'], wo['qty_to_build'])
        db.commit()

    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    cursor.execute("""SELECT wi.*, i.item_code, i.desc as item_desc, i.mpn,
                             COALESCE(wi.manufacturer_name, m.mfr_short_name, m.mfr_full_name) as manufacturer_name,
                             u.unit_short_name as unit
                     FROM work_order_item wi
                     JOIN items i ON wi.item_id = i.item_id
                     LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                     LEFT JOIN units u ON i.unit_id = u.unit_id
                     WHERE wi.wo_id = %s AND COALESCE(wi.is_removed, 0) = 0
                     ORDER BY i.item_code""", (wo_id,))
    wo_items_raw = cursor.fetchall()

    wo_items = []
    for item in wo_items_raw:
        stock = get_item_stock(cursor, item['item_id'])
        cursor.execute("""
            SELECT COALESCE(SUM(qty_blocked), 0) as wo_blocked
            FROM stock_blocking
            WHERE item_id = %s AND wo_id = %s AND block_status = 0
        """, (item['item_id'], wo_id))
        item['wo_blocked'] = float(cursor.fetchone()['wo_blocked'] or 0)
        item['available'] = stock['available']
        item['blocked'] = stock['blocked']
        item['shortage'] = max(0.0, round(float(item['qty_required']) - item['available'] - item['wo_blocked'], 3))
        
        # Fetch individual active blocks for this item in this WO
        cursor.execute("""
            SELECT stock_block_id, CAST(qty_blocked AS DOUBLE) as qty_blocked, remarks
            FROM stock_blocking
            WHERE item_id = %s AND wo_id = %s AND block_status = 0
        """, (item['item_id'], wo_id))
        item['active_blocks'] = cursor.fetchall()
        
        wo_items.append(item)

    # Fetch removed items for the bottom section
    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    cursor.execute("""SELECT wi.*, i.item_code, i.desc as item_desc, i.mpn,
                             COALESCE(wi.manufacturer_name, m.mfr_short_name, m.mfr_full_name) as manufacturer_name,
                             u.unit_short_name as unit
                      FROM work_order_item wi
                      JOIN items i ON wi.item_id = i.item_id
                      LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                      LEFT JOIN units u ON i.unit_id = u.unit_id
                      WHERE wi.wo_id = %s AND wi.is_removed = 1
                      ORDER BY i.item_code""", (wo_id,))
    removed_items_raw = cursor.fetchall()
    
    removed_items = []
    for item in removed_items_raw:
        stock = get_item_stock(cursor, item['item_id'])
        item['available'] = stock['available']
        removed_items.append(item)

    # Fetch all items for item code edit dropdown
    cursor.execute("SELECT item_id, item_code, `desc` FROM items WHERE is_deleted=0 ORDER BY item_code")
    all_items = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('work_order_detail.html',
                           wo=wo,
                           wo_items=wo_items,
                           removed_items=removed_items,
                           status_map=STATUS_MAP,
                           all_items=all_items)

# ─── WO ITEM UPDATE (bom_qty edit → recalc required qty) ─
@app.route('/work-order/<int:wo_id>/item/<int:wi_id>/update', methods=['POST'])
def wo_item_update(wo_id, wi_id):
    bom_qty = float(request.form['bom_qty'])
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("SELECT qty_to_build FROM work_order WHERE wo_id=%s", (wo_id,))
    wo = cursor.fetchone()
    qty_required = round(bom_qty * float(wo['qty_to_build']), 3)

    cursor.execute("""UPDATE work_order_item 
                     SET bom_qty=%s, qty_required=%s 
                     WHERE wo_item_id=%s AND wo_id=%s""",
                   (bom_qty, qty_required, wi_id, wo_id))

    cursor.execute("SELECT item_id FROM work_order_item WHERE wo_item_id=%s", (wi_id,))
    wi_row = cursor.fetchone()
    item_id = wi_row['item_id'] if wi_row else None

    cursor.execute("""
        SELECT COALESCE(SUM(qty_blocked), 0) as wo_blocked
        FROM stock_blocking
        WHERE item_id = %s
        AND wo_id=%s AND block_status=0
    """, (item_id, wo_id))
    wo_blocked = float(cursor.fetchone()['wo_blocked'] or 0)

    # Get item stock to calculate shortage
    stock = get_item_stock(cursor, item_id) if item_id else {'available': 0.0}
    shortage = max(0.0, round(qty_required - stock['available'] - wo_blocked, 3))

    db.commit()
    cursor.close()
    db.close()

    needs_block_alert = wo_blocked < qty_required
    block_diff = round(qty_required - wo_blocked, 3)

    return jsonify({
        'success': True,
        'qty_required': qty_required,
        'wo_blocked': wo_blocked,
        'needs_block_alert': needs_block_alert,
        'block_diff': block_diff if needs_block_alert else 0,
        'shortage': shortage
    })

# ─── WO ITEM CHANGE ITEM CODE ───────────────────────────
@app.route('/work-order/<int:wo_id>/item/<int:wi_id>/change-item', methods=['POST'])
def wo_item_change(wo_id, wi_id):
    new_item_id = request.form['new_item_id']
    db = get_db()
    cursor = db.cursor()
    cursor.execute("""UPDATE work_order_item SET item_id=%s 
                     WHERE wo_item_id=%s AND wo_id=%s""",
                   (new_item_id, wi_id, wo_id))
    db.commit()
    cursor.close()
    db.close()
    return redirect(url_for('work_order_detail', wo_id=wo_id))

# ─── CALCULATE REQUIREMENTS (Recalculate from BOM) ──────
@app.route('/work-order/<int:wo_id>/calculate', methods=['POST'])
def work_order_calculate(wo_id):            #function not used at all(?)
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT * FROM work_order WHERE wo_id = %s", (wo_id,))
    wo = cursor.fetchone()
    if not wo:
        cursor.close()
        db.close()
        return "Work order not found", 404
    if not wo['bom_id']:
        cursor.close()
        db.close()
        flash('This work order has no BOM assigned, so requirements cannot be recalculated from a BOM.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    calculate_requirements_from_bom(cursor, wo_id, wo['bom_id'], wo['qty_to_build'])
    db.commit()
    cursor.close()
    db.close()
    flash('Requirements recalculated from BOM.', 'success')
    return redirect(url_for('work_order_detail', wo_id=wo_id))

# ─── ADD NEW ITEM TO REQUIREMENTS TABLE ─────────────────
@app.route('/work-order/<int:wo_id>/item/add', methods=['POST'])
def wo_item_add(wo_id):
    item_code = request.form.get('item_code', '').strip()
    if item_code.startswith("'"):
        item_code = item_code[1:]
    item_desc = request.form.get('item_desc', '').strip()
    mpn = request.form.get('mpn', '').strip()
    manufacturer_name = request.form.get('manufacturer_name', '').strip()
    bom_qty = request.form.get('bom_qty', '').strip()

    if not item_code or not bom_qty:
        flash('Item code and BOM qty are required.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    try:
        bom_qty = float(bom_qty)
    except ValueError:
        flash('BOM qty must be a number.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("SELECT qty_to_build FROM work_order WHERE wo_id = %s", (wo_id,))
    wo = cursor.fetchone()
    if not wo:
        cursor.close()
        db.close()
        return "Work order not found", 404
    qty_required = round(bom_qty * float(wo['qty_to_build']), 3)

    # Find existing item by code, otherwise create a new master item record
    cursor.execute("SELECT item_id FROM items WHERE item_code = %s AND is_deleted = 0", (item_code,))
    item = cursor.fetchone()
    if item:
        item_id = item['item_id']
    else:
        # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
        mfr_id_val = None
        if manufacturer_name:
            cursor.execute("SELECT mfr_id FROM manufacturer WHERE LOWER(mfr_short_name) = %s OR LOWER(mfr_full_name) = %s LIMIT 1", (manufacturer_name.lower(), manufacturer_name.lower()))
            m_found = cursor.fetchone()
            if m_found:
                mfr_id_val = m_found['mfr_id']
            else:
                cursor.execute("INSERT INTO manufacturer (mfr_short_name, mfr_full_name) VALUES (%s, %s)", (manufacturer_name, manufacturer_name))
                mfr_id_val = cursor.lastrowid

        cursor.execute("""INSERT INTO items (item_code, `desc`, mpn, mfr_id, is_deleted)
                         VALUES (%s, %s, %s, %s, 0)""",
                       (item_code, item_desc or None, mpn or None, mfr_id_val))
        item_id = cursor.lastrowid

    # PRIOR CODE COMMENTED OUT:
    # cursor.execute("""INSERT INTO work_order_item
    #                  (wo_id, item_id, bom_qty, qty_required, qty_issued, qty_returned, manufacturer_name)
    #                  VALUES (%s, %s, %s, %s, 0, 0, %s)
    #                  ON DUPLICATE KEY UPDATE
    #                      bom_qty = VALUES(bom_qty),
    #                      qty_required = VALUES(qty_required),
    #                      manufacturer_name = VALUES(manufacturer_name)""",
    #                (wo_id, item_id, bom_qty, qty_required, manufacturer_name or None))

    # ADDED: Check if item already exists to prevent duplicates
    cursor.execute("SELECT wo_item_id FROM work_order_item WHERE wo_id = %s AND item_id = %s", (wo_id, item_id))
    existing = cursor.fetchone()
    if existing:
        cursor.execute("""UPDATE work_order_item 
                          SET bom_qty = %s, qty_required = %s, is_removed = 0,
                              manufacturer_name = COALESCE(%s, manufacturer_name)
                          WHERE wo_item_id = %s""",
                       (bom_qty, qty_required, manufacturer_name or None, existing['wo_item_id']))
    else:
        cursor.execute("""INSERT INTO work_order_item
                         (wo_id, item_id, bom_qty, qty_required, qty_issued, qty_returned, manufacturer_name)
                         VALUES (%s, %s, %s, %s, 0, 0, %s)""",
                       (wo_id, item_id, bom_qty, qty_required, manufacturer_name or None))

    db.commit()
    cursor.close()
    db.close()
    flash(f'Item {item_code} added to requirements.', 'success')
    return redirect(url_for('work_order_detail', wo_id=wo_id))


# ─── REMOVE ITEMS FROM REQUIREMENTS (Soft Delete) ────────
@app.route('/work-order/<int:wo_id>/items/remove', methods=['POST'])
def remove_work_order_items(wo_id):
    wo_item_ids = request.form.getlist('wo_item_ids')
    if not wo_item_ids:
        flash('No items selected for removal.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))
    
    db = get_db()
    cursor = db.cursor()
    try:
        ids = [int(x) for x in wo_item_ids]
    except ValueError:
        cursor.close()
        db.close()
        flash('Invalid items selected.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))
        
    placeholders = ', '.join(['%s'] * len(ids))
    
    # Retrieve the item_ids associated with these work_order_items
    cursor.execute(f"""SELECT DISTINCT item_id FROM work_order_item 
                      WHERE wo_id = %s AND wo_item_id IN ({placeholders})""",
                   (wo_id, *ids))
    item_rows = cursor.fetchall()
    item_ids = [row[0] for row in item_rows]
    
    # Release any active stock blocks for these items in this work order
    if item_ids:
        placeholders_items = ', '.join(['%s'] * len(item_ids))
        cursor.execute(f"""UPDATE stock_blocking 
                           SET block_status = 1 
                           WHERE wo_id = %s AND item_id IN ({placeholders_items}) AND block_status = 0""",
                       (wo_id, *item_ids))
        
    cursor.execute(f"""UPDATE work_order_item 
                      SET is_removed = 1 
                      WHERE wo_id = %s AND wo_item_id IN ({placeholders})""",
                   (wo_id, *ids))
    db.commit()
    cursor.close()
    db.close()
    flash(f'Removed {len(ids)} item(s) from the work order requirements and released blocked stock back into storage.', 'success')
    return redirect(url_for('work_order_detail', wo_id=wo_id))


# ─── RESTORE ITEM BACK TO WORK ORDER ────────────────────
@app.route('/work-order/<int:wo_id>/item/<int:wi_id>/restore', methods=['POST'])
def restore_work_order_item(wo_id, wi_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute("""UPDATE work_order_item 
                      SET is_removed = 0 
                      WHERE wo_id = %s AND wo_item_id = %s""",
                   (wo_id, wi_id))
    db.commit()
    cursor.close()
    db.close()
    flash('Item restored back to the requirements.', 'success')
    return redirect(url_for('work_order_detail', wo_id=wo_id))

# ─── WORK ORDER EDIT ────────────────────────────────────
@app.route('/work-order/<int:wo_id>/edit', methods=['GET', 'POST'])
def work_order_edit(wo_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    if request.method == 'POST':
        description = request.form['description']
        status = request.form['status']
        remarks = request.form['remarks']
        cancelled_by = request.form.get('cancelled_by', '')
        qty_to_build = request.form['qty_to_build']

        # Get current bom_id before update
        cursor.execute("SELECT bom_id FROM work_order WHERE wo_id = %s", (wo_id,))
        wo_old = cursor.fetchone()
        bom_id = wo_old['bom_id'] if wo_old else None

        if int(status) == 4 and cancelled_by:
            cursor.execute("""UPDATE work_order SET description=%s, status=%s,
                             remarks=%s, cancelled_by=%s, cancelled_at=NOW(), qty_to_build=%s
                             WHERE wo_id=%s""",
                           (description, status, remarks, cancelled_by, qty_to_build, wo_id))
        else:
            cursor.execute("""UPDATE work_order SET description=%s, status=%s,
                             remarks=%s, qty_to_build=%s WHERE wo_id=%s""",
                           (description, status, remarks, qty_to_build, wo_id))

        # Recalculate requirements based on the new number of sets
        calculate_requirements_from_bom(cursor, wo_id, bom_id, qty_to_build)

        db.commit()
        cursor.close()
        db.close()
        return redirect(url_for('work_order_detail', wo_id=wo_id))
    cursor.execute("""SELECT wo.*, b.bom_code, b.description as bom_desc
                     FROM work_order wo
                     JOIN bom b ON wo.bom_id = b.bom_id
                     WHERE wo.wo_id = %s""", (wo_id,))
    wo = cursor.fetchone()
    cursor.close()
    db.close()
    return render_template('work_order_edit.html', wo=wo, status_map=STATUS_MAP)

# ─── STOCK BLOCK ADD ────────────────────────────────────
@app.route('/work-order/<int:wo_id>/block', methods=['POST'])
def stock_block_add(wo_id):
    item_id = request.form['item_id']
    qty_blocked = float(request.form['qty_blocked'])
    remarks = request.form.get('remarks', '')
    db = get_db()
    cursor = db.cursor(dictionary=True)

    stock = get_item_stock(cursor, item_id)
    if qty_blocked > stock['available']:
        cursor.close()
        db.close()
        flash(f"Cannot block {qty_blocked} — only {stock['available']} available in stock.", 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    cursor.execute("""INSERT INTO stock_blocking
                     (wo_id, item_id, qty_blocked, qty_issued_from_block, block_status, remarks)
                     VALUES (%s, %s, %s, 0, 0, %s)""",
                   (wo_id, item_id, qty_blocked, remarks))
    db.commit()
    cursor.close()
    db.close()
    return redirect(url_for('work_order_detail', wo_id=wo_id))

# ─── STOCK BLOCK RELEASE ────────────────────────────────
@app.route('/work-order/<int:wo_id>/block/<int:block_id>/release', methods=['POST'])
def stock_block_release(wo_id, block_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute("UPDATE stock_blocking SET block_status=1 WHERE stock_block_id=%s", (block_id,))
    db.commit()
    cursor.close()
    db.close()
    return redirect(url_for('work_order_detail', wo_id=wo_id))


# ─── STOCK BLOCK RELEASE QUANTITY ───────────────────────
@app.route('/work-order/<int:wo_id>/item/<int:item_id>/release', methods=['POST'])
def stock_block_release_qty(wo_id, item_id):
    try:
        qty_to_release = float(request.form['qty_to_release'])
    except ValueError:
        flash('Invalid quantity to release.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    if qty_to_release <= 0:
        flash('Quantity to release must be greater than zero.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    db = get_db()
    cursor = db.cursor(dictionary=True)

    # Calculate current blocked qty
    cursor.execute("""
        SELECT COALESCE(SUM(qty_blocked), 0) as wo_blocked
        FROM stock_blocking
        WHERE item_id = %s AND wo_id = %s AND block_status = 0
    """, (item_id, wo_id))
    current_blocked = float(cursor.fetchone()['wo_blocked'] or 0)

    if qty_to_release > current_blocked:
        cursor.close()
        db.close()
        flash(f'Cannot release {qty_to_release} — only {current_blocked} units are blocked.', 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    # Fetch active blocks ordered by ID (FIFO order)
    cursor.execute("""
        SELECT stock_block_id, qty_blocked FROM stock_blocking
        WHERE item_id = %s AND wo_id = %s AND block_status = 0
        ORDER BY stock_block_id ASC
    """, (item_id, wo_id))
    active_blocks = cursor.fetchall()

    remaining = qty_to_release
    for block in active_blocks:
        if remaining <= 0:
            break
        block_qty = float(block['qty_blocked'])
        if block_qty <= remaining:
            # Release this block completely
            cursor.execute("UPDATE stock_blocking SET block_status=1 WHERE stock_block_id=%s", (block['stock_block_id'],))
            remaining -= block_qty
        else:
            # Partially release this block
            new_qty = round(block_qty - remaining, 3)
            cursor.execute("UPDATE stock_blocking SET qty_blocked=%s WHERE stock_block_id=%s", (new_qty, block['stock_block_id']))
            remaining = 0.0

    db.commit()
    cursor.close()
    db.close()
    flash(f'Successfully released {qty_to_release} units.', 'success')
    return redirect(url_for('work_order_detail', wo_id=wo_id))

# ─── ISSUE STOCK ────────────────────────────────────────
@app.route('/work-order/<int:wo_id>/issue', methods=['POST'])
def issue_stock(wo_id):
    item_id = int(request.form['item_id'])
    qty_to_issue = float(request.form['qty_to_issue'])
    remarks = request.form.get('remarks', '')
    db = get_db()
    cursor = db.cursor(dictionary=True)

    stock = get_item_stock(cursor, item_id)
    cursor.execute("""
        SELECT COALESCE(SUM(qty_blocked), 0) as wo_blocked
        FROM stock_blocking WHERE item_id=%s AND wo_id=%s AND block_status=0
    """, (item_id, wo_id))
    wo_blocked = float(cursor.fetchone()['wo_blocked'] or 0)
    real_available = stock['available'] + wo_blocked

    if qty_to_issue > real_available:
        cursor.close()
        db.close()
        flash(f"Cannot issue {qty_to_issue} — only {real_available} available.", 'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    #cursor.execute("""UPDATE stock_blocking SET block_status=1
     #                WHERE item_id=%s AND wo_id=%s AND block_status=0""",
      #             (item_id, wo_id))
        # Release blocks FIFO, but only up to the quantity actually being issued —
    # any remainder stays blocked/reserved for this WO
    remaining_to_release = qty_to_issue
    cursor.execute("""
        SELECT stock_block_id, qty_blocked FROM stock_blocking
        WHERE item_id=%s AND wo_id=%s AND block_status=0
        ORDER BY stock_block_id ASC
    """, (item_id, wo_id))
    active_blocks = cursor.fetchall()

    for block in active_blocks:
        if remaining_to_release <= 0:
            break
        block_qty = float(block['qty_blocked'])
        if block_qty <= remaining_to_release:
            cursor.execute("UPDATE stock_blocking SET block_status=1 WHERE stock_block_id=%s",
                           (block['stock_block_id'],))
            remaining_to_release -= block_qty
        else:
            new_qty = round(block_qty - remaining_to_release, 3)
            cursor.execute("UPDATE stock_blocking SET qty_blocked=%s WHERE stock_block_id=%s",
                           (new_qty, block['stock_block_id']))
            remaining_to_release = 0.0
    cursor.execute("SELECT wo_number, description FROM work_order WHERE wo_id = %s", (wo_id,))
    wo_row = cursor.fetchone()
    if wo_row and wo_row.get('description'):
        issued_to_str = f"WO #{wo_row['wo_number']} — {wo_row['description']}"
    elif wo_row:
        issued_to_str = f"WO #{wo_row['wo_number']}"
    else:
        issued_to_str = f"WO #{wo_id}"

    issued_by_str = session.get('username') or session.get('full_name') or 'admin'

    cursor.execute("""INSERT INTO issue_transaction 
                     (issue_date, item_id, wo_id, qty_issued, issued_to, issued_by, remarks, created_at)
                     VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())""",
                   (date.today(), item_id, wo_id, qty_to_issue, issued_to_str, issued_by_str, remarks))

    cursor.execute("""UPDATE work_order_item SET qty_issued = qty_issued + %s
                     WHERE wo_id=%s AND item_id=%s""",
                   (qty_to_issue, wo_id, item_id))

    cursor.execute("""UPDATE storage SET physical_availability = GREATEST(0, physical_availability - %s)
                      WHERE item_id = %s""",
                   (qty_to_issue, item_id))

    db.commit()
    cursor.close()
    db.close()
    flash(f'Successfully issued {qty_to_issue} units.', 'success')
    return redirect(url_for('work_order_detail', wo_id=wo_id))

# ─── GRN LIST ───────────────────────────────────────────
@app.route('/grn')
def grn_list():
    search = request.args.get('search', '')
    type_filter = request.args.get('grn_type', '')
    item_search = request.args.get('item_search', '').strip()
    item_id_param = request.args.get('item_id', '').strip()

    db = get_db()
    cursor = db.cursor(dictionary=True)

    selected_item = None
    item_id_filter = None

    if item_id_param:
        try:
            item_id_filter = int(item_id_param)
        except ValueError:
            item_id_filter = None

    if item_id_filter:
        cursor.execute("""
            SELECT i.*, u.unit_short_name as unit
            FROM items i
            LEFT JOIN units u ON i.unit_id = u.unit_id
            WHERE i.item_id = %s AND i.is_deleted = 0
        """, (item_id_filter,))
        selected_item = cursor.fetchone()

    if not selected_item and item_search:
        cursor.execute("""
            SELECT i.*, u.unit_short_name as unit
            FROM items i
            LEFT JOIN units u ON i.unit_id = u.unit_id
            WHERE i.is_deleted = 0 AND (i.item_code = %s OR i.item_code LIKE %s OR i.`desc` LIKE %s)
            ORDER BY CASE WHEN i.item_code = %s THEN 1 WHEN i.item_code LIKE %s THEN 2 ELSE 3 END, i.item_code
            LIMIT 1
        """, (item_search, f'%{item_search}%', f'%{item_search}%', item_search, f'{item_search}%'))
        selected_item = cursor.fetchone()
        if selected_item:
            item_id_filter = selected_item['item_id']

    query = """
        SELECT g.*, v.short_name AS vendor_name
        FROM grn g
        LEFT JOIN vendor v ON g.vendor_id = v.vendor_id
        WHERE 1=1
    """
    params = []
    if search:
        query += " AND (g.grn_no LIKE %s OR g.invoice_no LIKE %s)"
        params += [f'%{search}%', f'%{search}%']
    if type_filter != '':
        query += " AND g.grn_type = %s"
        params.append(int(type_filter))
    
    if item_id_filter:
        query += " AND g.grn_id IN (SELECT DISTINCT grn_id FROM grn_item WHERE item_id = %s)"
        params.append(item_id_filter)
    elif item_search and not selected_item:
        query += """ AND g.grn_id IN (
            SELECT DISTINCT gi.grn_id
            FROM grn_item gi
            JOIN items i ON gi.item_id = i.item_id
            WHERE i.item_code LIKE %s OR i.`desc` LIKE %s OR i.mpn LIKE %s
        )"""
        params += [f'%{item_search}%', f'%{item_search}%', f'%{item_search}%']

    query += " ORDER BY g.grn_id DESC"
    cursor.execute(query, params)
    grns = cursor.fetchall()

    item_stats = None
    if selected_item:
        cursor.execute("""
            SELECT 
                COALESCE(SUM(gi.qty_received), 0) AS total_received,
                COALESCE(SUM(gi.qty_accepted), 0) AS total_accepted,
                COALESCE(SUM(gi.qty_rejected), 0) AS total_rejected
            FROM grn_item gi
            JOIN grn g ON gi.grn_id = g.grn_id
            WHERE gi.item_id = %s
        """, (selected_item['item_id'],))
        st_row = cursor.fetchone()
        if st_row:
            tot_rec = float(st_row['total_received'] or 0)
            tot_acc = float(st_row['total_accepted'] or 0)
            tot_rej = float(st_row['total_rejected'] or 0)
            tot_pend = max(0.0, tot_rec - tot_acc - tot_rej)
            item_stats = {
                'total_received': tot_rec,
                'total_accepted': tot_acc,
                'total_pending': tot_pend,
                'total_rejected': tot_rej
            }

    for g in grns:
        cursor.execute("SELECT * FROM grn_item WHERE grn_id = %s", (g['grn_id'],))
        grn_items = cursor.fetchall()
        g['item_count'] = len(grn_items)
        g['grn_type_label'] = GRN_TYPE_MAP.get(g['grn_type'], 'Purchase')

        if selected_item:
            spec_qty = sum(float(it['qty_received'] or 0) for it in grn_items if it['item_id'] == selected_item['item_id'])
            g['specific_item_qty'] = spec_qty

        # 1. Determine GRN Status based on type & quantities received/ordered/returned
        grn_type = g.get('grn_type')
        if grn_type == 0:  # Purchase GRN
            po_id = g.get('po_id')
            total_ordered = 0
            if po_id:
                cursor.execute("SELECT SUM(qty_ordered) AS total_ord FROM po_items WHERE po_id = %s", (po_id,))
                res = cursor.fetchone()
                if res and res['total_ord']:
                    total_ordered = float(res['total_ord'])
            
            total_received = sum(float(it['qty_received'] or 0) for it in grn_items)
            
            if total_ordered > 0 and total_received < total_ordered:
                g['grn_status'] = 'Partially received'
                g['status_badge'] = 'badge-inprogress'
            else:
                g['grn_status'] = 'Received all'
                g['status_badge'] = 'badge-active'
                
        elif grn_type == 1:  # Work Order Return GRN
            wo_id = g.get('wo_id')
            total_wo_items = 0
            if wo_id:
                cursor.execute("SELECT COUNT(*) AS cnt FROM work_order_item WHERE wo_id = %s AND (is_removed IS NULL OR is_removed = 0)", (wo_id,))
                res = cursor.fetchone()
                if res:
                    total_wo_items = res['cnt']
            
            returned_item_count = len([it for it in grn_items if float(it['qty_received'] or 0) > 0])
            
            if total_wo_items > 0 and returned_item_count < total_wo_items:
                g['grn_status'] = 'Partial WO returned'
                g['status_badge'] = 'badge-inprogress'
            else:
                g['grn_status'] = 'Fully returned'
                g['status_badge'] = 'badge-active'
        else:
            # Internal Return, Online Purchase, Any Type, Other
            g['grn_status'] = '—'
            g['status_badge'] = ''
            
        # 2. Determine QC Status (QC Incomplete, QC In Progress, QC Complete)
        total_acc = sum(float(it['qty_accepted'] or 0) for it in grn_items)
        total_rej = sum(float(it['qty_rejected'] or 0) for it in grn_items)
        
        has_pending = False
        for it in grn_items:
            recv = float(it['qty_received'] or 0)
            acc = float(it['qty_accepted'] or 0)
            rej = float(it['qty_rejected'] or 0)
            if (recv - acc - rej) > 0:
                has_pending = True
                break
                
        if len(grn_items) == 0:
            g['qc_status'] = 'QC Incomplete'
            g['qc_status_badge'] = 'badge-cancelled'
        elif not has_pending:
            g['qc_status'] = 'QC Complete'
            g['qc_status_badge'] = 'badge-active'
        elif total_acc == 0 and total_rej == 0:
            g['qc_status'] = 'QC Incomplete'
            g['qc_status_badge'] = 'badge-cancelled'
        else:
            g['qc_status'] = 'QC In Progress'
            g['qc_status_badge'] = 'badge-inprogress'
            
    cursor.execute("""
        SELECT po.po_id, po.po_number, po.date_raised, po.po_version_number, v.short_name AS vendor_short_name, v.full_name AS vendor_full_name
        FROM purchase_order po
        LEFT JOIN vendor v ON po.vendor_id = v.vendor_id
        ORDER BY po.po_id DESC
    """)
    purchase_orders = cursor.fetchall()
    for po in purchase_orders:
        po['formatted_po_number'] = format_po_number(po['po_number'], po['date_raised'], po['po_version_number'], cursor=cursor)

    cursor.close()
    db.close()
    return render_template('grn_list.html', grns=grns, search=search, type_filter=type_filter, purchase_orders=purchase_orders, selected_item=selected_item, item_search=item_search, item_stats=item_stats)

# ─── GRN ADD (Step 1: GRN No & Type) ──────────────────────────
@app.route('/grn/add', methods=['GET', 'POST'])
def grn_add():
    db = get_db()
    cursor = db.cursor(dictionary=True)
    if request.method == 'POST':
        grn_type_int = int(request.form['grn_type'])
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        cursor.close()
        db.close()
        
        if grn_type_int == 0:
            return redirect(url_for('grn_add_purchase', grn_no=grn_no))
        elif grn_type_int == 1:
            return redirect(url_for('grn_add_wo_return', grn_no=grn_no))
        elif grn_type_int == 2:
            return redirect(url_for('grn_add_internal_return', grn_no=grn_no))
        elif grn_type_int == 3:
            return redirect(url_for('grn_add_online_purchase', grn_no=grn_no))
        elif grn_type_int == 4:
            return redirect(url_for('grn_add_any_type', grn_no=grn_no))
        elif grn_type_int == 5:
            return redirect(url_for('grn_add_work_order_return', grn_no=grn_no))
        else:
            return redirect(url_for('grn_add_other', grn_type=grn_type_int, grn_no=grn_no))

    next_grn_no = generate_grn_number(cursor)
    cursor.close()
    db.close()
    return render_template('grn_add.html', next_grn_no=next_grn_no)


# ─── GRN ADD PURCHASE (Step 2 for Purchase GRN) ───────────────
@app.route('/grn/add/purchase', methods=['GET', 'POST'])
def grn_add_purchase():
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        po_id = request.form.get('po_id') or None
        vendor_id = request.form.get('vendor_id') or None
        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date') or str(date.today())
        remarks = request.form.get('remarks', '')
        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')

        cursor.execute("""
            INSERT INTO grn (grn_no, grn_type, po_id, received_date, received_by, invoice_no, invoice_date, vendor_id, remarks, status, created_at)
            VALUES (%s, 0, %s, %s, %s, %s, %s, %s, %s, 'Received all', NOW())
        """, (grn_no, po_id, received_date, received_by, invoice_no, invoice_date, vendor_id, remarks))
        
        grn_id = cursor.lastrowid
        
        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))

        db.commit()
        cursor.close()
        db.close()
        flash(f'GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("""
        SELECT po.po_id, po.po_number, po.po_version_number, po.po_status, v.vendor_id,
               v.short_name AS vendor_short_name, v.full_name AS vendor_full_name
        FROM purchase_order po
        LEFT JOIN vendor v ON po.vendor_id = v.vendor_id
        WHERE po.po_status = 'Approved'
        ORDER BY po.po_id DESC
    """)
    purchase_orders = cursor.fetchall()
    
    cursor.close()
    db.close()
    return render_template('grn_add_purchase.html', grn_no=grn_no, purchase_orders=purchase_orders, today=date.today())


# ─── GRN ADD INTERNAL RETURN (Step 2 for Internal Return) ──────────
@app.route('/grn/add/internal-return', methods=['GET', 'POST'])
def grn_add_internal_return():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        returned_by = request.form.get('returned_by', '')
        return_date = request.form.get('return_date') or str(date.today())
        remarks = request.form.get('remarks', '')

        cursor.execute("""
            INSERT INTO grn (grn_no, grn_type, received_date, received_by, remarks, status, created_at)
            VALUES (%s, 2, %s, %s, %s, 'Received Item(s)', NOW())
        """, (grn_no, return_date, returned_by, remarks))
        
        grn_id = cursor.lastrowid

        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))

        db.commit()
        cursor.close()
        db.close()
        flash(f'Internal Return GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    returned_by = session.get('username') or session.get('full_name') or 'admin'
    
    cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
    items = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('grn_add_internal_return.html', grn_no=grn_no, returned_by=returned_by, items=items, today=date.today())


# ─── GRN ADD WORK ORDER RETURN (Step 2 for WO Return) ──────────
@app.route('/grn/add/wo-return', methods=['GET', 'POST'])
def grn_add_wo_return():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        wo_id_ref = request.form.get('wo_id_ref') or None
        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date') or str(date.today())
        project_details = (request.form.get('project_details') or '').strip()
        user_remarks = (request.form.get('remarks') or '').strip()
        
        remarks_parts = []
        if project_details:
            remarks_parts.append(f"Project: {project_details}")
        if user_remarks:
            remarks_parts.append(user_remarks)
        remarks = " | ".join(remarks_parts)

        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        cursor.execute("""
            INSERT INTO grn (grn_no, grn_type, received_date, received_by, invoice_no, invoice_date, wo_id, remarks, status, created_at)
            VALUES (%s, 1, %s, %s, %s, %s, %s, %s, 'Fully returned', NOW())
        """, (grn_no, received_date, received_by, invoice_no, invoice_date, wo_id_ref, remarks))
        
        grn_id = cursor.lastrowid

        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))

        db.commit()
        cursor.close()
        db.close()
        flash(f'Work Order Return GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("""
        SELECT wo_id, wo_number, description FROM work_order ORDER BY wo_id DESC
    """)
    work_orders = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('grn_add_wo_return.html', grn_no=grn_no, work_orders=work_orders, today=date.today())


# ─── GRN ADD TYPE 5: WORK ORDER RETURN ─────────────────────────
@app.route('/grn/add/work-order-return', methods=['GET', 'POST'])
def grn_add_work_order_return():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        wo_id_ref = request.form.get('wo_id_ref') or None
        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date') or str(date.today())
        project_details = (request.form.get('project_details') or '').strip()
        user_remarks = (request.form.get('remarks') or '').strip()
        
        remarks_parts = []
        if project_details:
            remarks_parts.append(f"Project: {project_details}")
        if user_remarks:
            remarks_parts.append(user_remarks)
        remarks = " | ".join(remarks_parts)

        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        cursor.execute("""
            INSERT INTO grn (grn_no, grn_type, received_date, received_by, invoice_no, invoice_date, wo_id, remarks, status, created_at)
            VALUES (%s, 5, %s, %s, %s, %s, %s, %s, 'Received Item(s)', NOW())
        """, (grn_no, received_date, received_by, invoice_no, invoice_date, wo_id_ref, remarks))
        
        grn_id = cursor.lastrowid

        bom_code = request.form.get('bom_code', '').strip()
        bom_desc = request.form.get('bom_desc', '').strip()
        qty_received_str = request.form.get('qty_received', '0').strip()
        line_remarks = request.form.get('line_remarks', '').strip()

        if bom_code:
            cursor.execute("SELECT item_id FROM items WHERE item_code = %s AND is_deleted = 0 LIMIT 1", (bom_code,))
            found_item = cursor.fetchone()
            if found_item:
                item_id = found_item['item_id']
            else:
                cursor.execute("""
                    INSERT INTO items (item_code, `desc`, is_deleted)
                    VALUES (%s, %s, 0)
                """, (bom_code, bom_desc or bom_code))
                item_id = cursor.lastrowid

            qty_r = float(qty_received_str or 0)
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, line_remarks))

        db.commit()
        cursor.close()
        db.close()
        flash(f'Work Order Return GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("""
        SELECT wo.wo_id, wo.wo_number, wo.description AS wo_description, b.bom_code, b.description AS bom_desc
        FROM work_order wo
        LEFT JOIN bom b ON wo.bom_id = b.bom_id
        ORDER BY wo.wo_id DESC
    """)
    work_orders = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('grn_add_work_order_return.html', grn_no=grn_no, work_orders=work_orders, today=date.today())


# ─── API WO ITEMS LOOKUP ─────────────────────────────────────────
@app.route('/api/wo-items/<int:wo_id>')
def api_wo_items(wo_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT wo_id, wo_number, description FROM work_order WHERE wo_id = %s", (wo_id,))
    wo = cursor.fetchone()
    if not wo:
        cursor.close()
        db.close()
        return jsonify({'success': False, 'message': 'Work Order not found'})

    cursor.execute("""
        SELECT woi.*, i.item_code, i.`desc` AS item_desc, i.mpn, u.unit_short_name AS unit
        FROM work_order_item woi
        JOIN items i ON woi.item_id = i.item_id
        LEFT JOIN units u ON i.unit_id = u.unit_id
        WHERE woi.wo_id = %s
        ORDER BY i.item_code
    """, (wo_id,))
    items = cursor.fetchall()
    cursor.close()
    db.close()

    return jsonify({
        'success': True,
        'wo': wo,
        'items': items
    })


# ─── GRN ADD ONLINE PURCHASE (Step 2 for Online Purchase) ──────
@app.route('/grn/add/online-purchase', methods=['GET', 'POST'])
def grn_add_online_purchase():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        po_number = request.form.get('po_number', '').strip()
        v_input = (request.form.get('vendor_id') or '').strip()
        vendor_id = int(v_input) if v_input and v_input.isdigit() and int(v_input) > 0 else None
        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date') or str(date.today())
        remarks = request.form.get('remarks', '')
        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        po_id = None
        if po_number:
            cursor.execute("SELECT po_id FROM purchase_order WHERE po_number LIKE %s LIMIT 1", (f"%{po_number}%",))
            res = cursor.fetchone()
            if res:
                po_id = res['po_id']

        cursor.execute("""
            INSERT INTO grn (grn_no, grn_type, po_number, po_id, received_date, received_by, invoice_no, invoice_date, vendor_id, remarks, status, created_at)
            VALUES (%s, 3, %s, %s, %s, %s, %s, %s, %s, %s, 'Received all', NOW())
        """, (grn_no, po_number, po_id, received_date, received_by, invoice_no, invoice_date, vendor_id, remarks))
        
        grn_id = cursor.lastrowid

        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))

        db.commit()
        cursor.close()
        db.close()
        flash(f'Online Purchase GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted = 0 ORDER BY short_name")
    vendors = cursor.fetchall()

    cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
    items = cursor.fetchall()

    cursor.execute("""
        SELECT po.po_id, po.po_number, po.date_raised, po.po_version_number, v.short_name AS vendor_short_name, v.full_name AS vendor_full_name
        FROM purchase_order po
        LEFT JOIN vendor v ON po.vendor_id = v.vendor_id
        ORDER BY po.po_id DESC
    """)
    purchase_orders = cursor.fetchall()
    for po in purchase_orders:
        po['formatted_po_number'] = format_po_number(po['po_number'], po['date_raised'], po['po_version_number'], cursor=cursor)

    cursor.close()
    db.close()
    return render_template('grn_add_online_purchase.html', grn_no=grn_no, vendors=vendors, items=items, purchase_orders=purchase_orders, today=date.today())


# ─── UPDATE GRN PO NUMBER ──────────────────────────────────────
@app.route('/grn/<int:grn_id>/edit-po', methods=['POST'])
def grn_edit_po(grn_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    po_number = request.form.get('po_number', '').strip()
    
    po_id = None
    if po_number:
        cursor.execute("SELECT po_id FROM purchase_order WHERE po_number LIKE %s LIMIT 1", (f"%{po_number}%",))
        res = cursor.fetchone()
        if res:
            po_id = res['po_id']

    cursor.execute("""
        UPDATE grn SET po_number = %s, po_id = %s WHERE grn_id = %s
    """, (po_number, po_id, grn_id))
    db.commit()
    cursor.close()
    db.close()
    flash(f'PO Number updated successfully.', 'success')
    return redirect(url_for('grn_list'))


# ─── GRN ADD ANY TYPE (Step 2 for Any Type) ───────────────────
@app.route('/grn/add/any-type', methods=['GET', 'POST'])
def grn_add_any_type():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        v_input = (request.form.get('vendor_id') or '').strip()
        vendor_id = int(v_input) if v_input and v_input.isdigit() and int(v_input) > 0 else None
        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date') or str(date.today())
        remarks = request.form.get('remarks', '')
        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        cursor.execute("""
            INSERT INTO grn (grn_no, grn_type, received_date, received_by, invoice_no, invoice_date, vendor_id, remarks, status, created_at)
            VALUES (%s, 4, %s, %s, %s, %s, %s, %s, 'Received all', NOW())
        """, (grn_no, received_date, received_by, invoice_no, invoice_date, vendor_id, remarks))
        
        grn_id = cursor.lastrowid

        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))

        db.commit()
        cursor.close()
        db.close()
        flash(f'Any Type GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted = 0 ORDER BY short_name")
    vendors = cursor.fetchall()

    cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
    items = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('grn_add_any_type.html', grn_no=grn_no, vendors=vendors, items=items, today=date.today())


# ─── CONVERT / MODIFY GRN TYPE ─────────────────────────────────
@app.route('/grn/<int:grn_id>/convert', methods=['GET', 'POST'])
def grn_convert(grn_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("SELECT * FROM grn WHERE grn_id = %s", (grn_id,))
    grn = cursor.fetchone()
    if not grn:
        cursor.close()
        db.close()
        return "GRN not found", 404

    if request.method == 'POST':
        target_type = int(request.form.get('target_type', 0))
        po_id = request.form.get('po_id') or None
        wo_id = request.form.get('wo_id_ref') or None
        po_number = request.form.get('po_number', '').strip()
        vendor_id = request.form.get('vendor_id') or grn.get('vendor_id')
        invoice_no = request.form.get('invoice_no') or grn.get('invoice_no')
        invoice_date = request.form.get('invoice_date') or grn.get('invoice_date') or str(date.today())
        
        project_details = (request.form.get('project_details') or '').strip()
        user_remarks = (request.form.get('remarks') or '').strip()
        remarks_parts = []
        if project_details:
            remarks_parts.append(f"Project: {project_details}")
        if user_remarks:
            remarks_parts.append(user_remarks)
        remarks = " | ".join(remarks_parts) if remarks_parts else grn.get('remarks')

        cursor.execute("""
            UPDATE grn 
            SET grn_type = %s, vendor_id = %s, invoice_no = %s, invoice_date = %s,
                wo_id = %s, po_number = %s, po_id = %s, remarks = %s, status = 'Received all'
            WHERE grn_id = %s
        """, (target_type, vendor_id, invoice_no, invoice_date, wo_id, po_number, po_id, remarks, grn_id))

        # Refresh line items
        cursor.execute("DELETE FROM grn_item WHERE grn_id = %s", (grn_id,))

        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))

        db.commit()
        cursor.close()
        db.close()
        target_name = GRN_TYPE_MAP.get(target_type, 'GRN')
        flash(f'GRN {grn["grn_no"]} type modified to {target_name} successfully.', 'success')
        return redirect(url_for('grn_list'))

    target_type = int(request.args.get('target_type', 0))

    cursor.execute("""
        SELECT gi.*, i.item_code, i.`desc` AS item_desc, i.mpn, u.unit_short_name AS unit
        FROM grn_item gi
        JOIN items i ON gi.item_id = i.item_id
        LEFT JOIN units u ON i.unit_id = u.unit_id
        WHERE gi.grn_id = %s
    """, (grn_id,))
    existing_items = cursor.fetchall()

    if target_type == 0: # Purchase
        cursor.execute("""
            SELECT po.po_id, po.po_number, po.po_version_number, po.po_status, v.vendor_id,
                   v.short_name AS vendor_short_name, v.full_name AS vendor_full_name
            FROM purchase_order po
            LEFT JOIN vendor v ON po.vendor_id = v.vendor_id
            WHERE po.po_status = 'Approved'
            ORDER BY po.po_id DESC
        """)
        purchase_orders = cursor.fetchall()
        cursor.close()
        db.close()
        return render_template('grn_add_purchase.html', grn_no=grn['grn_no'], purchase_orders=purchase_orders, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type)

    elif target_type == 1: # Work Order Return
        cursor.execute("SELECT wo_id, wo_number, description FROM work_order ORDER BY wo_id DESC")
        work_orders = cursor.fetchall()
        cursor.close()
        db.close()
        return render_template('grn_add_wo_return.html', grn_no=grn['grn_no'], work_orders=work_orders, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type)

    elif target_type == 2: # Internal Return
        cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
        items = cursor.fetchall()
        returned_by = grn.get('received_by') or session.get('username') or 'admin'
        cursor.close()
        db.close()
        return render_template('grn_add_internal_return.html', grn_no=grn['grn_no'], returned_by=returned_by, items=items, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type)

    elif target_type == 3: # Online Purchase
        cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted = 0 ORDER BY short_name")
        vendors = cursor.fetchall()
        cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
        items = cursor.fetchall()
        cursor.close()
        db.close()
        return render_template('grn_add_online_purchase.html', grn_no=grn['grn_no'], vendors=vendors, items=items, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type)

    elif target_type == 5: # Work Order Return
        cursor.execute("""
            SELECT wo.wo_id, wo.wo_number, wo.description AS wo_description, b.bom_code, b.description AS bom_desc
            FROM work_order wo
            LEFT JOIN bom b ON wo.bom_id = b.bom_id
            ORDER BY wo.wo_id DESC
        """)
        work_orders = cursor.fetchall()
        cursor.close()
        db.close()
        return render_template('grn_add_work_order_return.html', grn_no=grn['grn_no'], work_orders=work_orders, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type)

    cursor.close()
    db.close()
    return redirect(url_for('grn_list'))


# ─── GRN ADD OTHER (For Others) ──
@app.route('/grn/add/other', methods=['GET', 'POST'])
def grn_add_other():
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    if request.method == 'POST':
        grn_type_int = int(request.form.get('grn_type', 1))
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        wo_id_ref = request.form.get('wo_id_ref') or None
        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date') or str(date.today())
        v_input = (request.form.get('vendor_id') or '').strip()
        vendor_id = int(v_input) if v_input and v_input.isdigit() and int(v_input) > 0 else None
        remarks = request.form.get('remarks', '')
        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        cursor.execute("""
            INSERT INTO grn (grn_no, grn_type, received_date, received_by, invoice_no, invoice_date, vendor_id, wo_id, remarks, status, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, 'Received all', NOW())
        """, (grn_no, grn_type_int, received_date, received_by, invoice_no, invoice_date, vendor_id, wo_id_ref, remarks))
        
        grn_id = cursor.lastrowid
        
        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')
        
        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))

        db.commit()
        cursor.close()
        db.close()
        flash(f'GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_type_int = int(request.args.get('grn_type', 1))
    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    grn_type_label = GRN_TYPE_MAP.get(grn_type_int, 'Return')

    cursor.execute("SELECT wo_id, wo_number, description FROM work_order WHERE status IN (0,1,2,5) ORDER BY wo_id DESC")
    open_wos = cursor.fetchall()
    cursor.execute("SELECT item_id, item_code, `desc` as item_desc FROM items WHERE is_deleted=0 ORDER BY item_code")
    items = cursor.fetchall()
    cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted=0 ORDER BY short_name")
    vendors = cursor.fetchall()
    
    cursor.close()
    db.close()
    return render_template('grn_add_other.html', grn_no=grn_no, grn_type=grn_type_int, grn_type_label=grn_type_label, open_wos=open_wos, items=items, vendors=vendors, today=date.today())


# ─── API PO ITEMS LOOKUP ─────────────────────────────────────────
@app.route('/api/po-items/<int:po_id>')
def api_po_items(po_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("""
        SELECT po.*, v.vendor_id, v.short_name AS vendor_short_name, v.full_name AS vendor_full_name
        FROM purchase_order po
        LEFT JOIN vendor v ON po.vendor_id = v.vendor_id
        WHERE po.po_id = %s
    """, (po_id,))
    po = cursor.fetchone()
    if not po:
        cursor.close()
        db.close()
        return jsonify({'success': False, 'message': 'PO not found'})
    
    cursor.execute("""
        SELECT poi.*, i.item_code, i.`desc` AS item_desc, u.unit_short_name AS unit
        FROM po_items poi
        JOIN items i ON poi.item_id = i.item_id
        LEFT JOIN units u ON poi.unit_id = u.unit_id
        WHERE poi.po_id = %s
    """, (po_id,))
    items = cursor.fetchall()
    cursor.close()
    db.close()
    
    vendor_name = (po.get('vendor_short_name') + ' — ' + po.get('vendor_full_name')) if po.get('vendor_short_name') and po.get('vendor_full_name') else (po.get('vendor_full_name') or po.get('vendor_short_name') or 'N/A')
    return jsonify({
        'success': True,
        'po': po,
        'vendor_id': po.get('vendor_id'),
        'vendor_name': vendor_name,
        'items': items
    })


# ─── GRN DETAIL ─────────────────────────────────────────
@app.route('/grn/<int:grn_id>')
def grn_detail(grn_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("""
        SELECT g.*, v.short_name AS vendor_short_name, v.full_name AS vendor_full_name
        FROM grn g
        LEFT JOIN vendor v ON g.vendor_id = v.vendor_id
        WHERE g.grn_id = %s
    """, (grn_id,))
    grn = cursor.fetchone()
    if not grn:
        return "GRN not found", 404
    grn['grn_type_label'] = GRN_TYPE_MAP.get(grn['grn_type'], 'Purchase')
    
    # Map status badge for header
    status = grn.get('status') or 'Received'
    if status in ['Received Item(s)', 'Accepted']:
        grn['status_badge'] = 'badge-active'
    elif status in ['Item(s) Pending', 'Partially Accepted', 'In Progress']:
        grn['status_badge'] = 'badge-inprogress'
    elif status in ['Rejected', 'Cancelled']:
        grn['status_badge'] = 'badge-cancelled'
    else:
        grn['status_badge'] = 'badge-open'

    cursor.execute("""SELECT gi.*, i.item_code, i.`desc` as item_desc, i.mpn,
                             u.unit_short_name as unit
                     FROM grn_item gi
                     JOIN items i ON gi.item_id = i.item_id
                     LEFT JOIN units u ON i.unit_id = u.unit_id
                     WHERE gi.grn_id = %s ORDER BY gi.grn_item_id""", (grn_id,))
    grn_items = cursor.fetchall()
    wo_ref = None
    if grn.get('wo_id'):
        cursor.execute("SELECT wo_number, description FROM work_order WHERE wo_id = %s", (grn['wo_id'],))
        wo_ref = cursor.fetchone()
    cursor.close()
    db.close()
    return render_template('grn_detail.html', grn=grn, grn_items=grn_items, wo_ref=wo_ref)

# ─── GRN QC INSPECTION ──────────────────────────────────
@app.route('/grn/<int:grn_id>/qc', methods=['GET', 'POST'])
def grn_qc(grn_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    if request.method == 'POST':
        inspected_by = request.form.get('inspected_by') or session.get('username') or 'admin'
        inspection_date = request.form.get('inspection_date') or date.today()
        
        grn_item_ids = request.form.getlist('grn_item_id[]')
        qty_accepted = request.form.getlist('qty_accepted[]')
        qty_rejected = request.form.getlist('qty_rejected[]')
        qc_statuses = request.form.getlist('qc_status[]')
        line_remarks = request.form.getlist('line_remarks[]')
        
        for i, gi_id in enumerate(grn_item_ids):
            q_acc = float(qty_accepted[i] or 0)
            q_rej = float(qty_rejected[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            
            cursor.execute("SELECT item_id, location, qty_received, posted_qty_accepted FROM grn_item WHERE grn_item_id = %s", (gi_id,))
            existing_gi = cursor.fetchone()
            if existing_gi:
                item_id = existing_gi['item_id']
                loc = existing_gi['location']
                q_recv = float(existing_gi['qty_received'] or 0)
                prev_posted = float(existing_gi['posted_qty_accepted'] or 0)
                
                # Auto-calculate item qc_status
                pend = max(0.0, q_recv - q_acc - q_rej)
                if pend == 0:
                    if q_rej == 0:
                        status = 'Accepted'
                    elif q_acc == 0:
                        status = 'Rejected'
                    else:
                        status = 'Partially Accepted'
                else:
                    if q_acc > 0 or q_rej > 0:
                        status = 'In Progress'
                    else:
                        status = 'Pending'
                
                delta = q_acc - prev_posted
                if delta != 0:
                    s_row = None
                    if loc:
                        cursor.execute("SELECT store_id, physical_availability FROM storage WHERE item_id = %s AND store_location = %s LIMIT 1", (item_id, loc))
                        s_row = cursor.fetchone()
                    if not s_row:
                        cursor.execute("SELECT store_id, physical_availability FROM storage WHERE item_id = %s ORDER BY store_id ASC LIMIT 1", (item_id,))
                        s_row = cursor.fetchone()
                        
                    if s_row:
                        cursor.execute("UPDATE storage SET physical_availability = physical_availability + %s WHERE store_id = %s", (delta, s_row['store_id']))
                    else:
                        store_loc = loc or 'Main Store'
                        cursor.execute("INSERT INTO storage (item_id, store_location, physical_availability, created_at) VALUES (%s, %s, %s, NOW())", (item_id, store_loc, delta))
                
                cursor.execute("""
                    UPDATE grn_item
                    SET qty_accepted = %s, qty_rejected = %s, posted_qty_accepted = %s, qc_status = %s, remarks = %s,
                        qc_approved_by = %s, qc_date = %s
                    WHERE grn_item_id = %s
                """, (q_acc, q_rej, q_acc, status, rem, inspected_by, inspection_date, gi_id))
            
        cursor.execute("SELECT qty_received, qty_accepted, qty_rejected FROM grn_item WHERE grn_id = %s", (grn_id,))
        all_items = cursor.fetchall()
        all_pend_zero = all((float(it['qty_received'] or 0) - float(it['qty_accepted'] or 0) - float(it['qty_rejected'] or 0)) <= 0 for it in all_items) if all_items else False
        
        if all_pend_zero:
            all_accepted = all(float(it['qty_rejected'] or 0) == 0 for it in all_items)
            all_rejected = all(float(it['qty_accepted'] or 0) == 0 for it in all_items)
            if all_accepted:
                qc_overall_status = 'Accepted'
            elif all_rejected:
                qc_overall_status = 'Rejected'
            else:
                qc_overall_status = 'Partially Accepted'
        else:
            qc_overall_status = 'In Progress'
            
        cursor.execute("""
            UPDATE grn
            SET status = %s, inspected_by = %s, inspection_date = %s
            WHERE grn_id = %s
        """, (qc_overall_status, inspected_by, inspection_date, grn_id))
        
        db.commit()
        cursor.close()
        db.close()
        flash('QC Inspection results updated successfully.', 'success')
        return redirect(url_for('grn_list'))
        
    cursor.execute("""
        SELECT g.*, v.short_name AS vendor_short_name, v.full_name AS vendor_full_name
        FROM grn g
        LEFT JOIN vendor v ON g.vendor_id = v.vendor_id
        WHERE g.grn_id = %s
    """, (grn_id,))
    grn = cursor.fetchone()
    if not grn:
        return "GRN not found", 404
        
    grn['grn_type_label'] = GRN_TYPE_MAP.get(grn['grn_type'], 'Purchase')
    
    # Map status badge for header
    status = grn.get('status') or 'Received'
    if status == 'Received':
        grn['status_badge'] = 'badge-open'
    elif status == 'Accepted':
        grn['status_badge'] = 'badge-active'
    elif status == 'Rejected':
        grn['status_badge'] = 'badge-cancelled'
    elif status == 'Partially Accepted':
        grn['status_badge'] = 'badge-inprogress'
    else:
        grn['status_badge'] = 'badge-open'

    cursor.execute("""
        SELECT gi.*, i.item_code, i.`desc` as item_desc
        FROM grn_item gi
        JOIN items i ON gi.item_id = i.item_id
        WHERE gi.grn_id = %s 
        ORDER BY gi.grn_item_id
    """, (grn_id,))
    grn_items = cursor.fetchall()
    
    logged_user = session.get('full_name') or session.get('username') or 'admin'
    today = date.today()
    
    cursor.close()
    db.close()
    return render_template('grn_qc.html', grn=grn, grn_items=grn_items, logged_user=logged_user, today=today)


# ─── API GRN QC DETAILS ─────────────────────────────────────────
@app.route('/api/grn/<int:grn_id>/qc-details')
def api_grn_qc_details(grn_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("""
        SELECT gi.*, i.item_code, i.`desc` as item_desc
        FROM grn_item gi
        JOIN items i ON gi.item_id = i.item_id
        WHERE gi.grn_id = %s
        ORDER BY gi.grn_item_id
    """, (grn_id,))
    items = cursor.fetchall()
    
    result = []
    for it in items:
        recv = float(it['qty_received'] or 0)
        acc = float(it['qty_accepted'] or 0)
        rej = float(it['qty_rejected'] or 0)
        pend = max(0.0, recv - acc - rej)
        
        if pend == 0 and rej == 0:
            st = 'Accepted'
        elif pend == 0 and acc == 0:
            st = 'Rejected'
        elif pend == 0:
            st = 'Partially Accepted'
        elif acc > 0 or rej > 0:
            st = 'In Progress'
        else:
            st = 'Pending'
            
        result.append({
            'grn_item_id': it['grn_item_id'],
            'item_code': it['item_code'],
            'item_desc': it['item_desc'] or '',
            'qty_received': round(recv, 3),
            'qty_accepted': round(acc, 3),
            'qty_pending': round(pend, 3),
            'qty_rejected': round(rej, 3),
            'qc_status': st,
            'remarks': it['remarks'] or ''
        })
        
    cursor.close()
    db.close()
    return jsonify({'items': result})


# ─── API ITEM PREDICTION SEARCH ────────────────────────────────
@app.route('/api/items/search')
def api_items_search():
    q = request.args.get('q', '').strip()
    grn_id = request.args.get('grn_id', '').strip()
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    if grn_id:
        query = """
            SELECT DISTINCT i.item_id, i.item_code, i.`desc` as item_desc, i.mpn, u.unit_short_name as unit,
                   gi.qty_received, gi.qc_status
            FROM grn_item gi
            JOIN items i ON gi.item_id = i.item_id
            LEFT JOIN units u ON i.unit_id = u.unit_id
            WHERE gi.grn_id = %s
        """
        params = [grn_id]
        if q:
            query += " AND (i.item_code LIKE %s OR i.`desc` LIKE %s OR i.mpn LIKE %s)"
            params.extend([f'%{q}%', f'%{q}%', f'%{q}%'])
        query += " ORDER BY i.item_code LIMIT 20"
        cursor.execute(query, params)
    else:
        query = """
            SELECT i.item_id, i.item_code, i.`desc` as item_desc, i.mpn, u.unit_short_name as unit
            FROM items i
            LEFT JOIN units u ON i.unit_id = u.unit_id
            WHERE i.is_deleted = 0
        """
        params = []
        if q:
            query += " AND (i.item_code LIKE %s OR i.`desc` LIKE %s OR i.mpn LIKE %s)"
            params.extend([f'%{q}%', f'%{q}%', f'%{q}%'])
        query += " ORDER BY i.item_code LIMIT 20"
        cursor.execute(query, params)
        
    items = cursor.fetchall()
    cursor.close()
    db.close()
    return jsonify({'items': items})


def parse_storage_file(filename, file_bytes):
    """Reads Excel (.xlsx, .xls) or CSV (.csv, .tsv, .txt) files and extracts storage rows.
    Uses intelligent header detection and fallback column matching so no valid data is missed.
    """
    ext = (filename or '').lower().rsplit('.', 1)[-1] if '.' in (filename or '') else ''
    raw_rows = []

    if ext in ('csv', 'tsv', 'txt') or 'csv' in (filename or '').lower():
        try:
            text = file_bytes.decode('utf-8-sig', errors='replace')
        except Exception:
            text = file_bytes.decode('latin-1', errors='replace')
        stream = io.StringIO(text)
        try:
            dialect = csv.Sniffer().sniff(text[:2048])
            reader = csv.reader(stream, dialect)
        except Exception:
            stream.seek(0)
            reader = csv.reader(stream)
        for r in reader:
            if any(cell and str(cell).strip() for cell in r):
                raw_rows.append([str(c).strip() for c in r])
    else:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if any(c is not None and str(c).strip() for c in row):
                    raw_rows.append([str(c).strip() if c is not None else '' for c in row])
        except Exception as ex:
            raise ValueError(f"Could not parse spreadsheet file. Please ensure it is a valid .xlsx or .csv file. Error: {ex}")

    if not raw_rows:
        return []

    def clean_key(val):
        if val is None:
            return ""
        s = str(val).lower().strip()
        return "".join(c if c.isalnum() else " " for c in s)

    code_aliases = [
        'item code', 'item_code', 'itemcode', 'item no', 'item_no', 'item number', 'item_number',
        'code', 'part number', 'part_number', 'part no', 'part_no', 'partno', 'part', 'part#', 'item#',
        'sku', 'material', 'material code', 'material_code', 'product code', 'component', 'article', 'pn', 'p/n'
    ]
    qty_aliases = [
        'physical store availability', 'physical_store_availability',
        'physical availability', 'physical_availability',
        'physical count', 'physical_count',
        'qty', 'quantity', 'count', 'stock', 'avail',
        'availability', 'available', 'store qty', 'store_qty', 'total qty', 'total_qty',
        'balance', 'on hand', 'on_hand', 'in stock', 'in_stock', 'units', 'inventory'
    ]
    loc_aliases = [
        'location', 'store location', 'store_location', 'bin', 'bin location', 'bin_location',
        'rack', 'store', 'storage', 'warehouse', 'sub store', 'substore'
    ]
    desc_aliases = [
        'desc', 'description', 'item description', 'item_description', 'part description',
        'part_description', 'details', 'name', 'item name', 'specification'
    ]
    mpn_aliases = [
        'mpn', 'mfr part', 'mfr part no', 'mfr_part_no', 'mfg part', 'manufacturer part',
        'manufacturer part number', 'model'
    ]
    manu_aliases = [
        'manufacturer', 'manu', 'make', 'brand', 'mfr', 'manufacturer name', 'mfr name'
    ]

    header_row_idx = 0
    best_score = -1
    max_scan = min(20, len(raw_rows))

    for r_idx in range(max_scan):
        row_cleaned = [clean_key(c) for c in raw_rows[r_idx]]
        row_joined = ' '.join(row_cleaned)

        score = 0
        has_code = any(alias in row_joined for alias in code_aliases)
        has_qty = any(alias in row_joined for alias in qty_aliases)

        if has_code:
            score += 3
        if has_qty:
            score += 3

        if score > best_score and (has_code or has_qty):
            best_score = score
            header_row_idx = r_idx

    headers = [clean_key(c) for c in raw_rows[header_row_idx]]

    def find_col(aliases):
        for i, h in enumerate(headers):
            if not h:
                continue
            for alias in aliases:
                if alias in h:
                    return i
        return None

    code_col = find_col(code_aliases)
    qty_col = find_col(qty_aliases)
    loc_col = find_col(loc_aliases)
    desc_col = find_col(desc_aliases)
    mpn_col = find_col(mpn_aliases)
    manu_col = find_col(manu_aliases)

    # Fallback column detection if headers were not explicitly matched
    if code_col is None:
        code_col = 0
    if qty_col is None:
        for c_idx in range(len(headers)):
            if c_idx == code_col:
                continue
            num_count = 0
            for r_idx in range(header_row_idx + 1, min(header_row_idx + 10, len(raw_rows))):
                val_str = raw_rows[r_idx][c_idx] if c_idx < len(raw_rows[r_idx]) else ''
                try:
                    float(val_str)
                    num_count += 1
                except ValueError:
                    pass
            if num_count >= 1:
                qty_col = c_idx
                break
        if qty_col is None and len(headers) > 1:
            qty_col = 1 if code_col != 1 else 0

    parsed_rows = []
    for r_idx in range(header_row_idx + 1, len(raw_rows)):
        row = raw_rows[r_idx]

        def get_val(col_idx):
            if col_idx is None or col_idx >= len(row):
                return None
            v = row[col_idx]
            return str(v).strip() if v is not None else None

        item_code = get_val(code_col)
        if not item_code or item_code == '':
            continue

        if clean_key(item_code) in code_aliases:
            continue

        qty_str = get_val(qty_col)
        location = get_val(loc_col)
        desc = get_val(desc_col)
        mpn = get_val(mpn_col)
        manufacturer = get_val(manu_col)

        parsed_rows.append({
            'row_num': r_idx + 1,
            'item_code': item_code,
            'qty': qty_str,
            'location': location,
            'desc': desc,
            'mpn': mpn,
            'manufacturer': manufacturer
        })
    return parsed_rows


# ─── STORAGE IMPORT TEMPLATE DOWNLOAD ─────────────────────
@app.route('/storage/template')
def storage_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Storage Import Template"

    headers = ["Item Code", "Description", "MPN", "Manufacturer", "Quantity", "Location"]
    ws.append(headers)

    header_fill = openpyxl.styles.PatternFill(start_color="1A2E4A", end_color="1A2E4A", fill_type="solid")
    header_font = openpyxl.styles.Font(color="FFFFFF", bold=True)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font

    ws.append(["ITEM-1001", "100K Resistor 0805", "RES-0805-100K", "Yageo", 500, "Main Store Bin A-1"])
    ws.append(["ITEM-1002", "10uF Capacitor 16V", "CAP-0805-10UF", "Murata", 250, "Main Store Bin B-3"])
    ws.append(["ITEM-1003", "Microcontroller STM32F4", "STM32F407VGT6", "STMicroelectronics", 50, "Cabinet C-2"])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="storage_import_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ─── STORAGE IMPORT ─────────────────────────────────────
@app.route('/storage/import', methods=['GET', 'POST'])
def storage_import():
    if request.method == 'POST':
        if 'excel_file' not in request.files:
            flash('No file uploaded. Please select an Excel or CSV file.', 'error')
            return redirect(request.url)
        file = request.files['excel_file']
        if not file or file.filename == '':
            flash('No file selected. Please choose a file to import.', 'error')
            return redirect(request.url)

        default_location = request.form.get('default_location', 'Main Store').strip()
        if not default_location:
            default_location = 'Main Store'

        try:
            file_bytes = file.read()
            rows = parse_storage_file(file.filename, file_bytes)

            if not rows:
                flash('No valid item rows found in the uploaded file. Please check file headers or download our sample template.', 'error')
                return redirect(request.url)

            db = get_db()
            cursor = db.cursor(dictionary=True)

            imported_count = 0
            skipped_items = []

            for row in rows:
                item_code = row['item_code']
                qty = row['qty']
                location = row['location'] or default_location
                row_num = row.get('row_num', '?')

                try:
                    qty_val = float(qty) if qty is not None else 0.0
                except (ValueError, TypeError):
                    qty_val = 0.0

                if qty_val < 0:
                    qty_val = 0.0

                # STRICT RULE: NEVER add anything to items table!
                # Only look up existing items in the items master table
                cursor.execute("SELECT item_id FROM items WHERE item_code = %s AND (is_deleted IS NULL OR is_deleted = 0) LIMIT 1", (item_code,))
                item_row = cursor.fetchone()

                if not item_row:
                    # Item does not exist in items master — DO NOT CREATE IT, SKIP AND NOTIFY
                    skipped_items.append({
                        'item_code': item_code,
                        'row': row_num,
                        'qty': qty_val
                    })
                    continue

                item_id = item_row['item_id']

                # Check if storage record already exists for item_id
                cursor.execute("SELECT store_id, physical_availability FROM storage WHERE item_id = %s LIMIT 1", (item_id,))
                existing_store = cursor.fetchone()
                if existing_store:
                    new_total = float(existing_store['physical_availability'] or 0) + qty_val
                    cursor.execute("UPDATE storage SET physical_availability = %s WHERE store_id = %s", (new_total, existing_store['store_id']))
                else:
                    cursor.execute(
                        """INSERT INTO storage (item_id, store_location, physical_availability, created_at)
                           VALUES (%s, %s, %s, NOW())""",
                        (item_id, location, qty_val)
                    )
                imported_count += 1

            db.commit()
            cursor.close()
            db.close()

            if imported_count > 0:
                flash(f"Successfully imported {imported_count} storage inventory records.", "success")

            if skipped_items:
                preview_list = [f"Row {item['row']}: '{item['item_code']}'" for item in skipped_items[:15]]
                details_text = ', '.join(preview_list)
                if len(skipped_items) > 15:
                    details_text += f" ... (and {len(skipped_items) - 15} more)"

                flash(
                    f"⚠️ NOTIFICATION: {len(skipped_items)} item(s) were SKIPPED because they DO NOT exist in the items master table:\n[{details_text}].\nNote: As per strict policy, NO NEW ITEMS WERE ADDED TO THE ITEMS MASTER TABLE.",
                    "warning"
                )

            return redirect(url_for('storage_list'))

        except Exception as e:
            flash(f"An error occurred during import: {str(e)}", "error")
            return redirect(request.url)

    return render_template('storage_import.html')


# ─── STORAGE LIST ───────────────────────────────────────
@app.route('/storage')
def storage_list():
    search = request.args.get('search', '')
    db = get_db()
    cursor = db.cursor(dictionary=True)
    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    query = """SELECT s.*, i.item_code, i.`desc` as item_desc, i.mpn, COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name
               FROM storage s
               JOIN items i ON s.item_id = i.item_id
               LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
               WHERE 1=1"""
    params = []
    if search:
        query += " AND (i.item_code LIKE %s OR s.store_location LIKE %s OR i.mpn LIKE %s OR m.mfr_short_name LIKE %s OR m.mfr_full_name LIKE %s)"
        params += [f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%']
    query += " ORDER BY s.store_id ASC"
    cursor.execute(query, params)
    storages = cursor.fetchall()
    
    total_physical_availability = 0.0
    # Calculate currently available quantities (after blocking)
    for s in storages:
        stock = get_item_stock(cursor, s['item_id'])
        s['currently_available'] = stock['available']
        total_physical_availability += float(s.get('physical_availability') or 0)
        
    if total_physical_availability.is_integer():
        total_physical_availability = int(total_physical_availability)

    cursor.close()
    db.close()
    return render_template('storage_list.html', storages=storages, search=search, total_physical_availability=total_physical_availability)

# ─── STORAGE ADD ────────────────────────────────────────
@app.route('/storage/add', methods=['GET', 'POST'])
def storage_add():
    db = get_db()
    cursor = db.cursor(dictionary=True)
    if request.method == 'POST':
        item_id = request.form['item_id']
        store_location = request.form['store_location']
        physical_availability = float(request.form.get('physical_availability', 0) or 0)
        
        cursor.execute("SELECT store_id, store_location, physical_availability FROM storage WHERE item_id = %s LIMIT 1", (item_id,))
        existing = cursor.fetchone()
        if existing:
            new_total = float(existing['physical_availability'] or 0) + physical_availability
            cursor.execute("UPDATE storage SET physical_availability = %s WHERE store_id = %s", (new_total, existing['store_id']))
            flash(f"Added {physical_availability} units to existing storage location '{existing['store_location']}' (New Total: {new_total}).", 'success')
        else:
            cursor.execute("""INSERT INTO storage (item_id, store_location, physical_availability, created_at)
                              VALUES (%s, %s, %s, NOW())""",
                           (item_id, store_location, physical_availability))
            flash('Storage entry added successfully.', 'success')
            
        db.commit()
        cursor.close()
        db.close()
        return redirect(url_for('storage_list'))

    cursor.execute("SELECT item_id, item_code, `desc` as item_desc FROM items WHERE is_deleted=0 ORDER BY item_code")
    items = cursor.fetchall()
    cursor.close()
    db.close()
    return render_template('storage_add.html', items=items)


# ─── API ITEM STORAGE LOCATION ─────────────────────────
@app.route('/api/item-storage-location/<int:item_id>')
def api_item_storage_location(item_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT store_location, physical_availability FROM storage WHERE item_id = %s LIMIT 1", (item_id,))
    row = cursor.fetchone()
    cursor.close()
    db.close()
    if row:
        return jsonify({
            'exists': True,
            'location': row['store_location'],
            'physical_availability': float(row['physical_availability'] or 0)
        })
    return jsonify({'exists': False})

# ─── STORAGE EDIT ───────────────────────────────────────
@app.route('/storage/<int:store_id>/edit', methods=['GET', 'POST'])
def storage_edit(store_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    if request.method == 'POST':
        item_id = request.form['item_id']
        store_location = request.form['store_location']
        physical_availability = request.form.get('physical_availability', 0)
        cursor.execute("""UPDATE storage SET item_id=%s, store_location=%s, physical_availability=%s
                          WHERE store_id=%s""",
                       (item_id, store_location, physical_availability, store_id))
        db.commit()
        cursor.close()
        db.close()
        flash('Storage entry updated successfully.', 'success')
        return redirect(url_for('storage_list'))

    cursor.execute("""SELECT s.*, i.item_code 
                      FROM storage s
                      LEFT JOIN items i ON s.item_id = i.item_id
                      WHERE s.store_id=%s""", (store_id,))
    storage = cursor.fetchone()
    if not storage:
        cursor.close()
        db.close()
        return "Storage entry not found", 404
    cursor.execute("SELECT item_id, item_code, `desc` as item_desc FROM items WHERE is_deleted=0 ORDER BY item_code")
    items = cursor.fetchall()
    cursor.close()
    db.close()
    return render_template('storage_edit.html', storage=storage, items=items)

# ─── STORAGE DELETE ─────────────────────────────────────
@app.route('/storage/<int:store_id>/delete', methods=['POST'])
def storage_delete(store_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute("DELETE FROM storage WHERE store_id=%s", (store_id,))
    db.commit()
    cursor.close()
    db.close()
    flash('Storage entry deleted.', 'success')
    return redirect(url_for('storage_list'))


# ─── HSN MASTER & MAPPING ──────────────────────────────
@app.route('/hsn')
def hsn_list():
    search = request.args.get('search', '').strip()
    db = get_db()
    cursor = db.cursor(dictionary=True)

    # Migrated from hsn.item_id — now sourced via hsn_items junction table
    query = """
        SELECT h.*, hi.item_id, i.item_code, i.desc as item_desc
        FROM hsn h
        LEFT JOIN hsn_items hi ON h.hsn_id = hi.hsn_id
        LEFT JOIN items i ON hi.item_id = i.item_id
        WHERE h.is_deleted = 0
    """
    params = []
    if search:
        query += " AND (h.hsn_code LIKE %s OR h.description LIKE %s OR i.item_code LIKE %s)"
        params += [f'%{search}%', f'%{search}%', f'%{search}%']
    query += " ORDER BY h.hsn_id DESC"
    cursor.execute(query, params)
    hsn_list = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('hsn_list.html', hsn_list=hsn_list, search=search)


@app.route('/hsn/add', methods=['GET', 'POST'])
def hsn_add():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        item_id = int(request.form['item_id'])
        hsn_code = request.form['hsn_code'].strip()
        description = request.form.get('description', '').strip()
        tax_rate = float(request.form.get('tax_rate', 0))
        is_igst = 'is_igst' in request.form
        tax_date_from = request.form.get('tax_date_from') or date.today()

        if is_igst:
            cgst = 0.0
            sgst = 0.0
            igst = tax_rate
        else:
            cgst = round(tax_rate / 2.0, 2)
            sgst = round(tax_rate / 2.0, 2)
            igst = 0.0

        try:
            # Migrated from hsn.item_id — now sourced via hsn_items junction table
            cursor.execute("""
                INSERT INTO hsn (hsn_code, description, tax_rate, cgst, sgst, igst, tax_date_from, is_deleted)
                VALUES (%s, %s, %s, %s, %s, %s, %s, 0)
            """, (hsn_code, description, tax_rate, cgst, sgst, igst, tax_date_from))
            hsn_id = cursor.lastrowid
            
            # Migrated from items.hsn_id — now sourced via hsn_items junction table
            cursor.execute("""
                INSERT INTO hsn_items (item_id, hsn_id)
                VALUES (%s, %s)
                ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)
            """, (item_id, hsn_id))
            db.commit()
            flash(f"HSN Code {hsn_code} created successfully for the selected item.", "success")
        except mysql.connector.Error as err:
            flash(f"Database error: {err}", "error")
        finally:
            cursor.close()
            db.close()
        return redirect(url_for('hsn_list'))

    # Migrated from items.hsn_id — now sourced via hsn_items junction table
    cursor.execute("""
        SELECT i.item_id, i.item_code, i.`desc`, hi.hsn_id 
        FROM items i 
        LEFT JOIN hsn_items hi ON i.item_id = hi.item_id 
        WHERE i.is_deleted = 0 
        ORDER BY i.item_code
    """)
    items = cursor.fetchall()
    cursor.close()
    db.close()
    return render_template('hsn_add.html', items=items, today=date.today())


@app.route('/hsn/<int:hsn_id>/edit', methods=['GET', 'POST'])
def hsn_edit(hsn_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    # Migrated from hsn.item_id — now sourced via hsn_items junction table
    cursor.execute("""
        SELECT h.*, hi.item_id 
        FROM hsn h 
        LEFT JOIN hsn_items hi ON h.hsn_id = hi.hsn_id 
        WHERE h.hsn_id = %s AND h.is_deleted = 0
    """, (hsn_id,))
    hsn = cursor.fetchone()
    if not hsn:
        cursor.close()
        db.close()
        flash("HSN entry not found.", "error")
        return redirect(url_for('hsn_list'))

    if request.method == 'POST':
        item_id = int(request.form['item_id'])
        hsn_code = request.form['hsn_code'].strip()
        description = request.form.get('description', '').strip()
        tax_rate = float(request.form.get('tax_rate', 0))
        is_igst = 'is_igst' in request.form
        tax_date_from = request.form.get('tax_date_from') or date.today()

        if is_igst:
            cgst = 0.0
            sgst = 0.0
            igst = tax_rate
        else:
            cgst = round(tax_rate / 2.0, 2)
            sgst = round(tax_rate / 2.0, 2)
            igst = 0.0

        try:
            # Check if item association changed
            # Migrated from hsn.item_id — now sourced via hsn_items junction table
            cursor.execute("SELECT item_id FROM hsn_items WHERE hsn_id = %s LIMIT 1", (hsn_id,))
            old_item_row = cursor.fetchone()
            old_item_id = old_item_row['item_id'] if old_item_row else None

            if old_item_id is not None and old_item_id != item_id:
                cursor.execute("DELETE FROM hsn_items WHERE item_id = %s", (old_item_id,))
            
            # Migrated from hsn.item_id — now sourced via hsn_items junction table
            cursor.execute("""
                UPDATE hsn
                SET hsn_code = %s, description = %s, tax_rate = %s, cgst = %s, sgst = %s, igst = %s, tax_date_from = %s
                WHERE hsn_id = %s
            """, (hsn_code, description, tax_rate, cgst, sgst, igst, tax_date_from, hsn_id))
            
            # Migrated from items.hsn_id — now sourced via hsn_items junction table
            cursor.execute("""
                INSERT INTO hsn_items (item_id, hsn_id)
                VALUES (%s, %s)
                ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)
            """, (item_id, hsn_id))
            
            db.commit()
            flash(f"HSN Code {hsn_code} updated successfully.", "success")
        except mysql.connector.Error as err:
            flash(f"Database error: {err}", "error")
        finally:
            cursor.close()
            db.close()
        return redirect(url_for('hsn_list'))

    # Migrated from items.hsn_id — now sourced via hsn_items junction table
    cursor.execute("""
        SELECT i.item_id, i.item_code, i.`desc`, hi.hsn_id 
        FROM items i 
        LEFT JOIN hsn_items hi ON i.item_id = hi.item_id 
        WHERE i.is_deleted = 0 
        ORDER BY i.item_code
    """)
    items = cursor.fetchall()
    cursor.close()
    db.close()
    return render_template('hsn_edit.html', hsn=hsn, items=items)


@app.route('/hsn/<int:hsn_id>/delete', methods=['POST'])
def hsn_delete(hsn_id):
    db = get_db()
    cursor = db.cursor()
    cursor.execute("UPDATE hsn SET is_deleted = 1 WHERE hsn_id = %s", (hsn_id,))
    # Migrated from items.hsn_id — now sourced via hsn_items junction table
    cursor.execute("DELETE FROM hsn_items WHERE hsn_id = %s", (hsn_id,))
    db.commit()
    cursor.close()
    db.close()
    flash("HSN entry deleted.", "success")
    return redirect(url_for('hsn_list'))


# ─── ISSUE TRANSACTION LIST ─────────────────────────────
@app.route('/issue-transaction', methods=['GET', 'POST'])
@app.route('/issue-transactions', methods=['GET', 'POST'])
@app.route('/issue-transaction-history', methods=['GET', 'POST'])
def issue_transaction_list():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        issue_date = request.form.get('issue_date') or str(date.today())
        item_id_str = request.form.get('item_id')
        qty_issued_str = request.form.get('qty_issued')
        issued_to = (request.form.get('issued_to') or '').strip()
        wo_id_str = request.form.get('wo_id')
        issued_by = (request.form.get('issued_by') or session.get('username') or session.get('full_name') or 'admin').strip()
        remarks = (request.form.get('remarks') or '').strip()

        if not item_id_str or not qty_issued_str:
            flash('Item and Quantity Issued are required.', 'error')
            return redirect(url_for('issue_transaction_list'))

        try:
            item_id = int(item_id_str)
            qty_issued = float(qty_issued_str)
            if qty_issued <= 0:
                raise ValueError
        except ValueError:
            flash('Quantity Issued must be a positive number.', 'error')
            return redirect(url_for('issue_transaction_list'))

        wo_id = int(wo_id_str) if wo_id_str and wo_id_str.isdigit() and int(wo_id_str) > 0 else None

        # Fetch item info for success message
        cursor.execute("SELECT item_code FROM items WHERE item_id = %s", (item_id,))
        item_row = cursor.fetchone()
        item_code = item_row['item_code'] if item_row else f"Item #{item_id}"

        # Insert issue transaction record
        cursor.execute("""
            INSERT INTO issue_transaction (issue_date, item_id, wo_id, qty_issued, issued_to, issued_by, remarks, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())
        """, (issue_date, item_id, wo_id, qty_issued, issued_to or None, issued_by, remarks or None))

        # If issued against a Work Order, update work_order_item.qty_issued
        if wo_id:
            cursor.execute("""
                UPDATE work_order_item 
                SET qty_issued = qty_issued + %s 
                WHERE wo_id = %s AND item_id = %s
            """, (qty_issued, wo_id, item_id))

        db.commit()
        cursor.close()
        db.close()

        flash(f"Manual Issue Transaction recorded successfully for '{item_code}' ({qty_issued} units).", 'success')
        return redirect(url_for('issue_transaction_list'))

    search = request.args.get('search', '').strip()

    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    query = """
        SELECT it.*, 
               i.item_code, 
               i.`desc` AS item_desc, 
               i.mpn, 
               COALESCE(m.mfr_short_name, m.mfr_full_name) AS manufacturer_name,
               u.unit_short_name AS unit,
               wo.wo_number
        FROM issue_transaction it
        JOIN items i ON it.item_id = i.item_id
        LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
        LEFT JOIN work_order wo ON it.wo_id = wo.wo_id
        LEFT JOIN units u ON i.unit_id = u.unit_id
        WHERE 1=1
    """
    params = []
    if search:
        query += """ AND (
            i.item_code LIKE %s OR 
            i.`desc` LIKE %s OR 
            i.mpn LIKE %s OR 
            it.issued_to LIKE %s OR 
            it.issued_by LIKE %s OR 
            wo.wo_number LIKE %s
        )"""
        params += [f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%']

    query += " ORDER BY it.issue_id DESC"
    cursor.execute(query, params)
    issues = cursor.fetchall()

    total_transactions = len(issues)
    total_qty_issued = sum(float(x.get('qty_issued') or 0) for x in issues)
    if total_qty_issued.is_integer():
        total_qty_issued = int(total_qty_issued)

    # Fetch active items list for the Add Issue form dropdown
    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    cursor.execute("""
        SELECT i.item_id, i.item_code, i.`desc`, i.mpn, COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name
        FROM items i
        LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
        WHERE i.is_deleted = 0
        ORDER BY i.item_code
    """)
    items_list = cursor.fetchall()

    # Fetch active work orders list for the WO selection dropdown
    cursor.execute("SELECT wo_id, wo_number FROM work_order WHERE status != 2 ORDER BY wo_id DESC")
    work_orders_list = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template(
        'issue_transaction.html',
        issues=issues,
        search=search,
        total_transactions=total_transactions,
        total_qty_issued=total_qty_issued,
        items_list=items_list,
        work_orders_list=work_orders_list,
        today=date.today()
    )


def seed_admin_user():
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="admin",
            database="tribi_db"
        )
        cursor = conn.cursor(dictionary=True)
        
        # Check if table exists, if not, create it
        try:
            cursor.execute("SELECT COUNT(*) as cnt FROM users")
            cursor.fetchall() # Consume result to avoid Unread Result error
        except mysql.connector.Error as err:
            if err.errno == 1146:  # Table doesn't exist
                print("Users table not found, creating it...")
                cursor.execute("""
                    CREATE TABLE users (
                        user_id INT AUTO_INCREMENT PRIMARY KEY,
                        username VARCHAR(50) NOT NULL UNIQUE,
                        password_hash VARCHAR(255) NOT NULL,
                        full_name VARCHAR(100),
                        role ENUM('admin', 'staff', 'approver') NOT NULL,
                        is_active TINYINT(1) NOT NULL DEFAULT 1,
                        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                        last_login DATETIME NULL
                    )
                """)
                conn.commit()
            else:
                raise err
        
        # Now check if empty and seed
        cursor.execute("SELECT COUNT(*) as cnt FROM users")
        row = cursor.fetchone()
        if row and row['cnt'] == 0:
            hashed = generate_password_hash('admin')
            cursor.execute("""
                INSERT INTO users (username, password_hash, full_name, role, is_active)
                VALUES (%s, %s, %s, %s, 1)
            """, ('admin', hashed, 'Administrator', 'admin'))
            conn.commit()
        # Ensure system_settings table exists
        try:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS system_settings (
                    setting_key VARCHAR(50) PRIMARY KEY,
                    setting_value TEXT
                )
            """)
            conn.commit()
        except Exception as err:
            print("System settings table creation error:", err)

        # Ensure issue_transaction table exists
        try:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS issue_transaction (
                    issue_id INT AUTO_INCREMENT PRIMARY KEY,
                    issue_date DATE NOT NULL,
                    item_id INT UNSIGNED NOT NULL,
                    wo_id INT NULL,
                    qty_issued DECIMAL(12, 4) NOT NULL DEFAULT 0.0000,
                    issued_to VARCHAR(255) NULL,
                    issued_by VARCHAR(100) NULL,
                    remarks TEXT NULL,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (item_id) REFERENCES items(item_id),
                    FOREIGN KEY (wo_id) REFERENCES work_order(wo_id)
                )
            """)
            conn.commit()
        except Exception as err:
            print("Issue transaction table creation error:", err)

        cursor.close()
        conn.close()
    except Exception as e:
        print("Seeding/Creation error:", e)


def seed_hsn():
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="admin",
            database="tribi_db"
        )
        cursor = conn.cursor(dictionary=True)
        
        # Soft delete dummy legacy 999999 HSN
        cursor.execute("UPDATE hsn SET is_deleted = 1 WHERE hsn_code = '999999'")
        conn.commit()

        # Standard HSN Master Definitions
        standard_hsns = [
            {'hsn_code': '853400', 'description': 'Printed Circuit Boards & Assemblies (PCBA)', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
            {'hsn_code': '854110', 'description': 'Diodes, Transistors & Semiconductor Devices', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
            {'hsn_code': '853321', 'description': 'Electrical Resistors, Capacitors & Passive Components', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
            {'hsn_code': '731815', 'description': 'Screws, Bolts, Nuts, Studs & Fasteners', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
            {'hsn_code': '854442', 'description': 'Insulated Wires, Cables, Harnesses & Connectors', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
            {'hsn_code': '761699', 'description': 'Aluminum Heatsinks & Mechanical Fittings', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
            {'hsn_code': '350691', 'description': 'Prepared Adhesives, Compounds & Coatings', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
            {'hsn_code': '854370', 'description': 'Electrical Apparatus & Electronic Assemblies', 'tax_rate': 18.00, 'cgst': 9.00, 'sgst': 9.00, 'igst': 18.00},
        ]
        
        hsn_map = {}
        for h in standard_hsns:
            cursor.execute("SELECT hsn_id FROM hsn WHERE hsn_code = %s AND is_deleted = 0 LIMIT 1", (h['hsn_code'],))
            row = cursor.fetchone()
            if row:
                hsn_map[h['hsn_code']] = row['hsn_id']
            else:
                cursor.execute("""
                    INSERT INTO hsn (hsn_code, description, tax_rate, cgst, sgst, igst, tax_date_from, is_deleted)
                    VALUES (%s, %s, %s, %s, %s, %s, '2020-01-01', 0)
                """, (h['hsn_code'], h['description'], h['tax_rate'], h['cgst'], h['sgst'], h['igst']))
                hsn_map[h['hsn_code']] = cursor.lastrowid
        conn.commit()

        # Migrated from items.hsn_id — now sourced via hsn_items junction table
        cursor.execute("""
            SELECT i.item_id, i.`desc` 
            FROM items i 
            LEFT JOIN hsn_items hi ON i.item_id = hi.item_id 
            LEFT JOIN hsn h ON hi.hsn_id = h.hsn_id 
            WHERE hi.hsn_id IS NULL OR h.hsn_code = '999999' OR h.is_deleted = 1
        """)
        unmapped_items = cursor.fetchall()
        
        def _categorize(desc):
            d = (desc or '').lower()
            if any(k in d for k in ['pcb', 'pcba', 'circuit board', 'mcd2000']):
                return '853400'
            elif any(k in d for k in ['diode', 'igbt', 'transistor', 'mosfet', 'rectifier', 'power module', 'semiconductor']):
                return '854110'
            elif any(k in d for k in ['res', 'capacitor', 'cap', 'inductor', 'transformer', 'thermistor', 'filter']):
                return '853321'
            elif any(k in d for k in ['screw', 'washer', 'stud', 'nut', 'spring', 'fastener']):
                return '731815'
            elif any(k in d for k in ['harness', 'cable', 'wire', 'connector', 'terminal', 'clamp', 'insulator', 'grommet']):
                return '854442'
            elif any(k in d for k in ['htsnk', 'heatsink', 'plate', 'enclosure', 'bracket', 'fitting']):
                return '761699'
            elif any(k in d for k in ['compound', 'coating', 'adhesive', 'glue', 'lacquer', 'tape']):
                return '350691'
            else:
                return '854370'

        for item in unmapped_items:
            code = _categorize(item['desc'])
            h_id = hsn_map.get(code)
            if h_id:
                # Migrated from items.hsn_id — now sourced via hsn_items junction table
                cursor.execute("""
                    INSERT INTO hsn_items (item_id, hsn_id)
                    VALUES (%s, %s)
                    ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)
                """, (item['item_id'], h_id))
        conn.commit()
            
        # Add is_locked column if not exists
        cursor.execute("SHOW COLUMNS FROM purchase_order LIKE 'is_locked'")
        if not cursor.fetchone():
            cursor.execute("ALTER TABLE purchase_order ADD COLUMN is_locked TINYINT(1) NOT NULL DEFAULT 0")
            conn.commit()
            print("Added is_locked column to purchase_order table.")

        # Drop unique key constraint on hsn_code if exists
        cursor.execute("SHOW INDEX FROM hsn WHERE Key_name = 'hsn_code'")
        if cursor.fetchone():
            cursor.execute("ALTER TABLE hsn DROP INDEX hsn_code")
            conn.commit()
            print("Dropped unique index hsn_code from hsn table.")

        # Add currency column to purchase_order if not exists
        cursor.execute("SHOW COLUMNS FROM purchase_order LIKE 'currency'")
        if not cursor.fetchone():
            cursor.execute("ALTER TABLE purchase_order ADD COLUMN currency VARCHAR(10) NOT NULL DEFAULT 'INR'")
            conn.commit()
            print("Added currency column to purchase_order table.")

        cursor.close()
        conn.close()
    except Exception as e:
        print("HSN seeding error:", e)


def format_po_number(po_number, date_raised, version, cursor=None):
    if date_raised:
        if isinstance(date_raised, str):
            from datetime import datetime
            try:
                dt = datetime.strptime(date_raised, '%Y-%m-%d').date()
            except ValueError:
                dt = date.today()
        else:
            dt = date_raised
    else:
        dt = date.today()
        
    fy = get_financial_year(cursor=cursor, date_obj=dt)
        
    seq = f"{po_number:04d}"
    
    try:
        ver_val = int(float(version)) if version is not None else 0
    except (ValueError, TypeError):
        ver_val = 0
        
    return f"{fy}/{seq}.{ver_val}"


@app.before_request
def require_login():
    allowed_endpoints = ['login', 'static']
    if request.path.startswith('/static/'):
        return
    if request.endpoint not in allowed_endpoints and 'user_id' not in session:
        return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
        
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password']
        
        db = get_db()
        cursor = db.cursor(dictionary=True)
        cursor.execute("SELECT * FROM users WHERE username = %s AND is_active = 1", (username,))
        user = cursor.fetchone()
        
        if user and check_password_hash(user['password_hash'], password):
            # Update last login timestamp
            from datetime import datetime
            cursor.execute("UPDATE users SET last_login = %s WHERE user_id = %s", (datetime.now(), user['user_id']))
            db.commit()
            
            # Store user details in session
            session['user_id'] = user['user_id']
            session['username'] = user['username']
            session['role'] = user['role']
            session['full_name'] = user['full_name']
            
            cursor.close()
            db.close()
            flash('Logged in successfully.', 'success')
            return redirect(url_for('dashboard'))
        else:
            cursor.close()
            db.close()
            flash('Invalid username or password.', 'error')
            return redirect(url_for('login'))
            
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    flash('Logged out successfully.', 'success')
    return redirect(url_for('login'))


@app.context_processor
def inject_global_policy_alert():
    try:
        policy_exp_date_str = get_setting('policy_expiration_date', '2026-12-31')
        policy_expired = False
        if policy_exp_date_str:
            exp_date = date.fromisoformat(policy_exp_date_str)
            if date.today() >= exp_date:
                policy_expired = True
    except Exception:
        policy_expired = False

    return {
        'policy_expired_alert': policy_expired,
        'policy_alert_message': 'The duration for the existing policy number has expired, please change the date/ policy condition.'
    }


@app.route('/settings')
def settings():
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT user_id, username, full_name, role, is_active, created_at, last_login FROM users ORDER BY username")
    users_list = cursor.fetchall()
    
    fy_settings = {
        'custom_enabled': get_setting('fy_custom_enabled', '0', cursor=cursor),
        'start_month': get_setting('fy_start_month', 'March', cursor=cursor),
        'start_year': get_setting('fy_start_year', '2027', cursor=cursor),
        'end_month': get_setting('fy_end_month', 'April', cursor=cursor),
        'end_year': get_setting('fy_end_year', '2028', cursor=cursor),
        'label': get_setting('fy_label', '2027-2028', cursor=cursor),
        'grn_format': get_setting('fy_grn_format', 'start_year', cursor=cursor),
        'current_fy_label': get_financial_year(cursor=cursor),
        'current_fy_start': get_financial_year_start(cursor=cursor)
    }

    policy_settings = {
        'line_text': get_setting('policy_line_text', 'COVERED UNDER NEW INDIA ASSURANCE POLICY NO. 67020021200200000030', cursor=cursor),
        'expiration_date': get_setting('policy_expiration_date', '2026-12-31', cursor=cursor)
    }
    
    cursor.close()
    db.close()
    return render_template('settings.html', users_list=users_list, fy=fy_settings, policy=policy_settings)


@app.route('/settings/financial-year', methods=['POST'])
def settings_financial_year():
    custom_enabled = request.form.get('custom_enabled', '0')
    start_month = request.form.get('start_month', 'March').strip()
    start_year = request.form.get('start_year', str(date.today().year)).strip()
    end_month = request.form.get('end_month', 'April').strip()
    end_year = request.form.get('end_year', str(date.today().year + 1)).strip()
    grn_format = request.form.get('grn_format', 'start_year').strip()
    
    if start_year and end_year:
        if start_year == end_year:
            fy_label = start_year
        else:
            fy_label = f"{start_year}-{end_year}"
    else:
        fy_label = start_year or end_year or str(date.today().year)

    db = get_db()
    cursor = db.cursor(dictionary=True)
    set_setting('fy_custom_enabled', '1' if custom_enabled == '1' else '0', cursor=cursor)
    set_setting('fy_start_month', start_month, cursor=cursor)
    set_setting('fy_start_year', start_year, cursor=cursor)
    set_setting('fy_end_month', end_month, cursor=cursor)
    set_setting('fy_end_year', end_year, cursor=cursor)
    set_setting('fy_label', fy_label, cursor=cursor)
    set_setting('fy_grn_format', grn_format, cursor=cursor)
    db.commit()
    cursor.close()
    db.close()

    flash(f"Financial Year settings updated successfully to {fy_label} ({start_month} {start_year} to {end_month} {end_year}).", "success")
    return redirect(url_for('settings'))


@app.route('/settings/policy', methods=['POST'])
def settings_policy():
    policy_line_text = (request.form.get('policy_line_text') or '').strip()
    policy_expiration_date = (request.form.get('policy_expiration_date') or '').strip()

    if not policy_line_text:
        policy_line_text = 'COVERED UNDER NEW INDIA ASSURANCE POLICY NO. 67020021200200000030'
    if not policy_expiration_date:
        policy_expiration_date = '2026-12-31'

    db = get_db()
    cursor = db.cursor(dictionary=True)
    set_setting('policy_line_text', policy_line_text, cursor=cursor)
    set_setting('policy_expiration_date', policy_expiration_date, cursor=cursor)
    db.commit()
    cursor.close()
    db.close()

    flash('Policy Number & Insurance settings updated successfully.', 'success')
    return redirect(url_for('settings'))


@app.route('/settings/user/add', methods=['POST'])
def settings_user_add():
    # Check permission (only admin can add users)
    if session.get('role') != 'admin':
        flash('Only administrators can create users.', 'error')
        return redirect(url_for('settings'))
        
    username = request.form['username'].strip()
    password = request.form['password']
    full_name = request.form['full_name'].strip()
    role = request.form['role']
    
    if not username or not password or not full_name or not role:
        flash('All fields are required.', 'error')
        return redirect(url_for('settings'))
        
    hashed = generate_password_hash(password)
    
    db = get_db()
    cursor = db.cursor()
    try:
        cursor.execute("""
            INSERT INTO users (username, password_hash, full_name, role, is_active)
            VALUES (%s, %s, %s, %s, 1)
        """, (username, hashed, full_name, role))
        db.commit()
        flash(f'User account for {username} created successfully.', 'success')
    except mysql.connector.Error as err:
        if err.errno == 1062:  # Duplicate entry
            flash('Username already exists. Please choose a different one.', 'error')
        else:
            flash(f'Database error: {err}', 'error')
    finally:
        cursor.close()
        db.close()
        
    return redirect(url_for('settings'))


@app.route('/settings/user/<int:user_id>/edit', methods=['POST'])
def settings_user_edit(user_id):
    if not session.get('user_id'):
        flash('Please log in to update user account.', 'error')
        return redirect(url_for('login'))

    curr_user_id = session.get('user_id')
    curr_role = session.get('role')

    # Strict access control check: Admin can edit any account; non-admin can ONLY edit their own account matching session user_id
    if curr_role != 'admin' and curr_user_id != user_id:
        flash('Access Denied: You can only edit details for your own account.', 'error')
        return redirect(url_for('settings'))

    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT * FROM users WHERE user_id = %s", (user_id,))
    target_user = cursor.fetchone()
    if not target_user:
        cursor.close()
        db.close()
        flash('User account not found.', 'error')
        return redirect(url_for('settings'))

    new_username = (request.form.get('username') or '').strip()
    new_full_name = (request.form.get('full_name') or '').strip()
    new_password = request.form.get('new_password') or ''
    new_role = request.form.get('role') or target_user['role']
    is_active_input = request.form.get('is_active')

    if not new_username or not new_full_name:
        cursor.close()
        db.close()
        flash('Username and Full Name are required.', 'error')
        return redirect(url_for('settings'))

    # Check if username is being changed and if new_username is taken by another user
    if new_username.lower() != target_user['username'].lower():
        cursor.execute("SELECT user_id FROM users WHERE LOWER(username) = %s AND user_id != %s", (new_username.lower(), user_id))
        if cursor.fetchone():
            cursor.close()
            db.close()
            flash('Username is already taken by another account.', 'error')
            return redirect(url_for('settings'))

    # Build SQL update query
    updates = ["username = %s", "full_name = %s"]
    params = [new_username, new_full_name]

    if new_password.strip():
        hashed = generate_password_hash(new_password)
        updates.append("password_hash = %s")
        params.append(hashed)

    # Only admin can change role or active status
    if curr_role == 'admin':
        updates.append("role = %s")
        params.append(new_role)
        if is_active_input is not None:
            updates.append("is_active = %s")
            params.append(1 if is_active_input == '1' else 0)
        else:
            # If admin unchecked active checkbox for another user (or self), set is_active = 0
            if 'is_active' in request.form or is_active_input is None:
                updates.append("is_active = %s")
                params.append(1 if is_active_input == '1' else 0)

    params.append(user_id)
    sql = f"UPDATE users SET {', '.join(updates)} WHERE user_id = %s"
    
    try:
        cursor.execute(sql, tuple(params))
        db.commit()

        # Update session details if current user modified their own account
        if curr_user_id == user_id:
            session['username'] = new_username
            session['full_name'] = new_full_name
            if curr_role == 'admin':
                session['role'] = new_role

        flash(f"User account for '{new_username}' updated successfully.", 'success')
    except Exception as err:
        flash(f"Error updating user account: {err}", 'error')
    finally:
        cursor.close()
        db.close()

    return redirect(url_for('settings'))


if __name__ == '__main__':
    seed_admin_user()
    seed_hsn()
    app.run(host='0.0.0.0', port=5000, debug=True, use_evalex=False)

#