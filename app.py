from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, send_file
import mysql.connector
from datetime import date, datetime
from flask import g  # to prevent database timeout
import openpyxl
import re
import os
import uuid
import secrets
import tempfile
import io
import csv
import json
import unicodedata
from fractions import Fraction
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
def _load_secret_key():
    """The key that signs session cookies.

    It was the literal string 'tribi_secret_key', sitting in the source. That
    is enough to forge a cookie: anyone who has seen this file can mint a
    session that says they are the admin, without a password, and the login
    screen never sees them.

    Order of preference:
      1. TRIBI_SECRET_KEY in the environment, for a real deployment.
      2. A .secret_key file beside this one, generated on first run.
      3. A fresh random key held in memory, if neither can be used — logins
         then stop working across a restart, which is visible and annoying
         rather than silent and dangerous.

    Changing the key logs everybody out once. That is the only effect.
    """
    key = os.environ.get('TRIBI_SECRET_KEY', '').strip()
    if len(key) >= 32:
        return key

    here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, '.secret_key')
    try:
        with open(path, 'r') as fh:
            key = fh.read().strip()
        if len(key) >= 32:
            return key
    except OSError:
        pass

    key = secrets.token_hex(32)
    try:
        with open(path, 'w') as fh:
            fh.write(key)
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass          # Windows; the file is still outside the source
    except OSError:
        print("WARNING: could not write .secret_key — sessions will not "
              "survive a restart.")
    return key


app.secret_key = _load_secret_key()

# Application version shown in the navbar. Update this one line on each release.
APP_VERSION = '1.0.a'

# Decimal places kept on every quantity.
#
# Three was enough while everything was counted in pieces. Copper wire is held
# by weight, and a board's worth of 34 SWG is 0.0004 kg — which rounds to zero
# at three places, so the requirement would simply vanish. Six places is a
# milligram, which covers every gauge in the store.
#
# The database columns must match: DECIMAL(14,6) on the quantity columns.
# Rounding here at six while the column holds three just moves the truncation
# somewhere less visible.
QTY_DP = 6

# Role restrictions are OFF for this deployment: every signed-in user can do
# everything, including reviewing, approving, cancelling, short closing and
# reopening a purchase order. That is a deliberate choice for the first
# version, not an oversight.
#
# The checks themselves are still in place and still read the role — they just
# consult this flag first. Turning restrictions back on is this one line, with
# nothing to find again and nothing to re-wire:
#
#     ENFORCE_ROLE_CHECKS = True
#
# Account creation is NOT covered by this. /settings/user/add stays admin-only,
# because handing every user the ability to mint an admin account is a
# different kind of permission from being allowed to approve an order.
ENFORCE_ROLE_CHECKS = False

# Whether a purchase order line must carry an HSN code. OFF for this
# deployment: a line may be saved without one, and the printed PO marks that
# line rather than inventing a rate for it.
#
# This needs po_items.hsn_id to be NULLable — see tribi_po_hsn_optional.sql.
# Without that ALTER the column cannot hold "no HSN" and the save fails on a
# NOT NULL constraint, which is why the flag and the migration go together.
#
# Turning it back on is this one line:
#     REQUIRE_HSN_ON_PO = True
REQUIRE_HSN_ON_PO = False

# Patterns for recognising 'Do Not Stuff' markings in bom_item.patch
_DNS_TOKEN_RE = re.compile(r'(^|[^a-z0-9])d[\s.\-_]*n[\s.\-_]*s([^a-z0-9]|$)', re.I)
_DO_NOT_STUFF_RE = re.compile(r'do\s*not\s*(stuff|populate|place|mount|fit)', re.I)

# ─── DATABASE CONNECTION SETTINGS ───────────────────────
# One place for the connection details, used by every connect() call below.
# MariaDB on port 3307; utf8mb4 so the rupee sign and other symbols store correctly.
DB_CONFIG = {
    'host': "localhost",
    'port': 3307,
    'user': "root",
    'password': "srigowri",
    'database': "tribi_db",
    'charset': "utf8mb4",
    'collation': "utf8mb4_unicode_ci",
}

_grn_schema_checked = False

def ensure_grn_schema():
    global _grn_schema_checked
    if _grn_schema_checked:
        return
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
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
    conn = mysql.connector.connect(**DB_CONFIG)
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

# PO short close reuses po_status (a VARCHAR, so no schema change needed).
SHORT_CLOSED_STATUS = 'Short Closed'

GRN_TYPE_MAP = {
    0: 'Purchase',
    1: 'Back To Store',
    2: 'Internal Return',
    3: 'Online Purchase',
    4: 'Miscellaneous',
    5: 'Work Order Output'
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


def is_dns_patch(patch):
    """True when a bom_item.patch cell means Do Not Stuff.

    The patch column carries two different kinds of note: DNS markings and
    ordinary assembly instructions ("For M3 Screw", "Put the sealant glue").
    Only the first kind excludes an item from a work order.

    Matches DNS as a standalone word, allowing spaces, dots, hyphens or
    underscores between the letters, so DNS / dns / d n s / D.N.S. / D-N-S
    all count, as do compound notes like "DNS - Lead Short" and "DNS-For Q4".
    Also matches spelled-out forms such as "do not stuff" / "do not populate".

    Deliberately NOT a plain substring test: real part numbers contain those
    letters (the ADNS-3080 optical sensor, for one), and excluding a genuine
    component from a build is a far worse failure than missing a DNS marking.
    """
    if not patch or not str(patch).strip():
        return False
    text = ' '.join(str(patch).split())
    return bool(_DNS_TOKEN_RE.search(text) or _DO_NOT_STUFF_RE.search(text))


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
        wi_unit_col = ('wi.bom_unit_id' if has_unit_column(cursor, 'work_order_item')
                       else 'NULL')
        cursor.execute(f"""SELECT wi.wo_item_id, wi.bom_qty, wi.item_id,
                                 {wi_unit_col} AS bom_unit_id, i.unit_id AS item_unit_id
                            FROM work_order_item wi
                            JOIN items i ON i.item_id = wi.item_id
                           WHERE wi.wo_id = %s""", (wo_id,))
        existing_items = cursor.fetchall()
        units = load_unit_map(cursor)
        conversions = load_item_conversions(cursor, [r['item_id'] for r in existing_items])
        for row in existing_items:
            bom_qty = float(row['bom_qty'] or 0)
            # The BOM quantity may be in its own unit. Scale by the build
            # quantity first, then convert — both are linear, but doing it in
            # this order keeps bom_qty on the row in the unit the BOM stated.
            per_build = bom_qty * float(qty_to_build)
            bom_unit = bom_unit_for(units, conversions, row['item_id'],
                                    row.get('bom_unit_id'), row.get('item_unit_id'))
            conv, status = bom_qty_to_stock(units, conversions, row['item_id'],
                                            per_build, bom_unit,
                                            row.get('item_unit_id'))
            if status in ('same', 'converted', 'item_factor'):
                cursor.execute("UPDATE work_order_item SET qty_required = %s WHERE wo_item_id = %s",
                               (round(conv, QTY_DP), row['wo_item_id']))
        return len(existing_items)

    # Migrated from items.manufacturer_name — now sourced via mfr_id -> manufacturer table
    bom_unit_col = 'bi.unit_id' if has_unit_column(cursor, 'bom_item') else 'NULL'
    cursor.execute(f"""SELECT bi.bom_item_id, bi.item_id, bi.qty, bi.patch, i.item_code,
                            {bom_unit_col} AS bom_unit_id, i.unit_id AS item_unit_id,
                            COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name
                     FROM bom_item bi
                     JOIN items i ON bi.item_id = i.item_id
                     LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                     WHERE bi.bom_id = %s AND bi.is_deleted = 0""", (bom_id,))
    all_bom_items = cursor.fetchall()

    # Do Not Stuff rows stay on the BOM (it must mirror the engineering sheet)
    # but never become work order requirements.
    bom_items = [bi for bi in all_bom_items if not is_dns_patch(bi.get('patch'))]
    dns_items = [bi for bi in all_bom_items if is_dns_patch(bi.get('patch'))]
    # Only report items that are excluded outright. An item with some DNS rows
    # and some normal rows still appears on the work order, so naming it as
    # 'excluded' would be wrong.
    kept_item_ids = {bi['item_id'] for bi in bom_items}
    fully_excluded = {}
    for d in dns_items:
        if d['item_id'] not in kept_item_ids:
            fully_excluded[d['item_id']] = {'item_code': d['item_code'], 'patch': d['patch']}
    calculate_requirements_from_bom.last_dns_excluded = list(fully_excluded.values())
    calculate_requirements_from_bom.last_dns_rows = len(dns_items)

    # Remove stale BOM-linked rows only (items that used to be in this BOM
    # but have since been removed from it). Anything with bom_item_id IS
    # NULL (manual 'Add new item' rows) is never touched. DNS items are kept
    # in the comparison list so recalculating an OLDER work order does not
    # delete DNS rows it already has — new work orders never get them anyway.
    # Match on item_id, not bom_item_id: after aggregation a work order row
    # represents ALL the BOM rows for that item, and only one of their ids is
    # stored. Comparing bom_item_ids would delete rows that are still valid.
    current_item_ids = list({bi['item_id'] for bi in all_bom_items})
    if current_item_ids:
        placeholders = ', '.join(['%s'] * len(current_item_ids))
        cursor.execute(f"""DELETE FROM work_order_item
                          WHERE wo_id = %s AND bom_item_id IS NOT NULL
                          AND item_id NOT IN ({placeholders})""",
                       (wo_id, *current_item_ids))
    else:
        cursor.execute("""DELETE FROM work_order_item
                         WHERE wo_id = %s AND bom_item_id IS NOT NULL""", (wo_id,))

    # An item can appear on a BOM many times — once per reference designator,
    # often across different sections. work_order_item holds ONE row per item,
    # so the per-row quantities must be SUMMED. Writing them one at a time
    # would leave only the last row's quantity, understating the requirement
    # (one real BOM lists the same part 19 times: 48 needed, 10 recorded).
    aggregated = {}
    for bi in bom_items:
        entry = aggregated.get(bi['item_id'])
        if entry is None:
            aggregated[bi['item_id']] = {
                'item_id': bi['item_id'],
                'bom_item_id': bi['bom_item_id'],      # first row represents the group
                'bom_qty': float(bi['qty'] or 0),
                'bom_unit_id': bi.get('bom_unit_id'),
                'item_unit_id': bi.get('item_unit_id'),
                'manufacturer_name': bi.get('manufacturer_name'),
                'times_listed': 1,
            }
        else:
            entry['bom_qty'] = round(entry['bom_qty'] + float(bi['qty'] or 0), QTY_DP)
            entry['times_listed'] += 1
            if not entry['manufacturer_name']:
                entry['manufacturer_name'] = bi.get('manufacturer_name')

    units = load_unit_map(cursor)
    conversions = load_item_conversions(cursor, list(aggregated.keys()))
    unconvertible = []

    for row in aggregated.values():
        # bom_qty stays in the unit the BOM stated; qty_required is what the
        # store must actually pull, in the unit it holds. Keeping both means a
        # later change to a conversion factor cannot silently restate history.
        per_build = row['bom_qty'] * float(qty_to_build)
        bom_unit = bom_unit_for(units, conversions, row['item_id'],
                                row.get('bom_unit_id'), row.get('item_unit_id'))
        conv, status = bom_qty_to_stock(units, conversions, row['item_id'],
                                        per_build, bom_unit,
                                        row.get('item_unit_id'))
        if status not in ('same', 'converted', 'item_factor'):
            # No honest figure exists. Record the line with no requirement
            # rather than one that is wrong by a factor of a thousand.
            unconvertible.append(row['item_id'])
            qty_required = None
        else:
            qty_required = round(conv, QTY_DP)

        cursor.execute("SELECT wo_item_id FROM work_order_item WHERE wo_id = %s AND item_id = %s",
                       (wo_id, row['item_id']))
        existing = cursor.fetchone()
        keep_unit = has_unit_column(cursor, 'work_order_item')
        if existing:
            if keep_unit:
                cursor.execute("""UPDATE work_order_item
                                  SET bom_item_id = %s, bom_qty = %s, bom_unit_id = %s,
                                      qty_required = %s,
                                      manufacturer_name = COALESCE(%s, manufacturer_name)
                                  WHERE wo_item_id = %s""",
                               (row['bom_item_id'], row['bom_qty'], row.get('bom_unit_id'),
                                qty_required, row['manufacturer_name'], existing['wo_item_id']))
            else:
                cursor.execute("""UPDATE work_order_item
                                  SET bom_item_id = %s, bom_qty = %s, qty_required = %s,
                                      manufacturer_name = COALESCE(%s, manufacturer_name)
                                  WHERE wo_item_id = %s""",
                               (row['bom_item_id'], row['bom_qty'], qty_required,
                                row['manufacturer_name'], existing['wo_item_id']))
        elif keep_unit:
            cursor.execute("""INSERT INTO work_order_item
                             (wo_id, item_id, bom_item_id, bom_qty, bom_unit_id, qty_required,
                              qty_issued, qty_returned, manufacturer_name)
                             VALUES (%s, %s, %s, %s, %s, %s, 0, 0, %s)""",
                           (wo_id, row['item_id'], row['bom_item_id'], row['bom_qty'],
                            row.get('bom_unit_id'), qty_required, row['manufacturer_name']))
        else:
            cursor.execute("""INSERT INTO work_order_item
                             (wo_id, item_id, bom_item_id, bom_qty, qty_required,
                              qty_issued, qty_returned, manufacturer_name)
                             VALUES (%s, %s, %s, %s, %s, 0, 0, %s)""",
                           (wo_id, row['item_id'], row['bom_item_id'], row['bom_qty'],
                            qty_required, row['manufacturer_name']))

    calculate_requirements_from_bom.last_unconvertible = unconvertible

    calculate_requirements_from_bom.last_merged = [
        {'item_id': r['item_id'], 'times_listed': r['times_listed'], 'bom_qty': r['bom_qty']}
        for r in aggregated.values() if r['times_listed'] > 1]
    return len(aggregated)


def get_table_columns(cursor, table_name):
    """Introspects live column metadata for a table in the current database.
    Used so we never hard-code the 'items' schema and never risk an
    ALTER-TABLE-style change — we only ever read structure and INSERT/UPDATE
    rows within columns that already exist."""
    cursor.execute("""
        SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE, COLUMN_DEFAULT, EXTRA,
               CHARACTER_MAXIMUM_LENGTH
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = %s
        ORDER BY ORDINAL_POSITION
    """, (table_name,))
    return cursor.fetchall()


_NUMERIC_TYPES = {'int', 'smallint', 'tinyint', 'mediumint', 'bigint',
                   'decimal', 'float', 'double'}


# Columns whose name says they are a ratio or multiplier. Zero would be a
# nonsense conversion factor and could divide by zero later, so these get 1.
_FACTOR_HINTS = ('conv_fact', 'conversion', 'factor', 'multiplier', 'ratio')


def _insert_with_defaults(cursor, table, values, overrides=None):
    """INSERT into `table`, filling in any other NOT NULL column that has no
    default of its own. The schema is read live, so a column added to the
    table later does not break the insert."""
    columns = get_table_columns(cursor, table)
    col_names = {c['COLUMN_NAME'] for c in columns}
    row = {k: v for k, v in values.items() if k in col_names}

    for col in columns:
        name = col['COLUMN_NAME']
        if name in row or 'auto_increment' in (col['EXTRA'] or ''):
            continue
        if overrides and name in overrides:
            if col['IS_NULLABLE'] == 'NO' and col['COLUMN_DEFAULT'] is None:
                row[name] = overrides[name]
            continue
        default = _safe_default_for_column(col)
        if default is not None:
            row[name] = default

    # Trim text to the column width. unit_short_name is varchar(10), so an
    # unusually long unit word from a spreadsheet would otherwise abort the
    # whole import under strict mode.
    widths = {c['COLUMN_NAME']: c.get('CHARACTER_MAXIMUM_LENGTH') for c in columns}
    for k, v in list(row.items()):
        w = widths.get(k)
        if w and isinstance(v, str) and len(v) > w:
            row[k] = v[:w]

    cols_sql = ', '.join(f"`{k}`" for k in row)
    placeholders = ', '.join(['%s'] * len(row))
    cursor.execute(f"INSERT INTO `{table}` ({cols_sql}) VALUES ({placeholders})",
                   list(row.values()))
    return cursor.lastrowid


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
        name = (col['COLUMN_NAME'] or '').lower()
        if any(h in name for h in _FACTOR_HINTS):
            return 1
        return 0
    if dtype == 'date':
        return date.today()
    if dtype in ('timestamp', 'datetime'):
        return None  # leave out; if truly NOT NULL with no default this is rare
    return ''  # varchar / text / etc.


# ─── UNIT CONVERSION ────────────────────────────────────
# A unit belongs to a family through base_unit_id: a base unit points at
# nothing, a child points at its base. The factor converts the child INTO the
# base, as an exact fraction:
#
#     qty_in_base = qty * conv_fact_num / conv_fact_denom
#
# Gram is 1/1000 of a kilogram, so 500 g = 0.5 kg. Converting between two units
# is therefore only possible when they share a base.

def load_unit_map(cursor):
    """Every unit keyed by unit_id. Small table, read once per request."""
    cursor.execute("""SELECT unit_id, unit_name, unit_short_name, base_unit_id,
                             conv_fact_num, conv_fact_denom, decimal_places
                      FROM units""")
    units = {}
    for r in cursor.fetchall():
        num = int(r['conv_fact_num'] or 1) or 1
        den = int(r['conv_fact_denom'] or 1) or 1
        units[int(r['unit_id'])] = {
            'unit_id': int(r['unit_id']),
            'name': r['unit_name'],
            'short': r['unit_short_name'] or r['unit_name'],
            'base_unit_id': int(r['base_unit_id']) if r['base_unit_id'] else None,
            'num': num,
            'denom': den,
            'decimals': r['decimal_places'] if r['decimal_places'] is not None else 3,
        }
    return units


def unit_base_id(units, unit_id):
    """The family this unit belongs to, named by its base unit's id.

    A base unit is its own family. The hop count is capped so a row that points
    at itself, or a pair that point at each other, cannot hang the request.
    """
    if unit_id is None:
        return None
    seen, current = set(), int(unit_id)
    for _ in range(10):
        u = units.get(current)
        if not u or not u['base_unit_id'] or current in seen:
            return current
        seen.add(current)
        current = u['base_unit_id']
    return current


def load_item_conversions(cursor, item_ids=None):
    """Per-item factors, keyed (item_id, alt_unit_id).

    These cross unit families, which the `units` table cannot express: there is
    no factor between kilograms and metres in general, only for one particular
    wire. Kept separate for exactly that reason.

    Returns {} if the table has not been created yet, so the app runs either
    side of the migration.
    """
    global _HAS_CONV_TABLE
    if _HAS_CONV_TABLE is False:
        return {}
    try:
        sql = """SELECT item_id, unit_id, stock_unit_id,
                        conv_fact_num, conv_fact_denom
                   FROM item_unit_conversion"""
        params = ()
        if item_ids:
            ids = [int(i) for i in item_ids if i]
            if not ids:
                return {}
            sql += " WHERE item_id IN (%s)" % ','.join(['%s'] * len(ids))
            params = tuple(ids)
        cursor.execute(sql, params)
    except Exception:
        _HAS_CONV_TABLE = False
        return {}
    _HAS_CONV_TABLE = True
    out = {}
    for r in cursor.fetchall():
        num = int(r['conv_fact_num'] or 0)
        den = int(r['conv_fact_denom'] or 0)
        if num <= 0 or den <= 0:
            continue          # a zero factor would collapse every quantity
        out[(int(r['item_id']), int(r['unit_id']))] = {
            'stock_unit_id': int(r['stock_unit_id']),
            'num': num,
            'denom': den,
        }
    return out


_HAS_CONV_TABLE = None


def _dim_factor(units, from_unit_id, to_unit_id):
    """The exact fraction converting one unit into another WITHIN a family.

    Returns None when they belong to different families, which is the signal to
    look for a per-item factor instead. Kept as a Fraction so a whole chain of
    hops can be multiplied out before touching a float even once.
    """
    if from_unit_id is None or to_unit_id is None:
        return None
    a, b = units.get(int(from_unit_id)), units.get(int(to_unit_id))
    if not a or not b:
        return None
    if a['unit_id'] == b['unit_id']:
        return Fraction(1)
    if unit_base_id(units, a['unit_id']) != unit_base_id(units, b['unit_id']):
        return None
    return Fraction(a['num'], a['denom']) / Fraction(b['num'], b['denom'])


def convert_qty(units, qty, from_unit_id, to_unit_id, conversions=None, item_id=None):
    """Converts a quantity between two units.

    Two kinds of conversion, tried in that order:

      1. DIMENSIONAL, from the `units` table. True of the universe - 1 kg is
         1000 g for everything - so it is preferred wherever it applies.
      2. PER ITEM, from `item_unit_conversion`. True of one part only: 1 kg of
         14 SWG copper is 34 m, 1 kg of 40 SWG is 9547 m.

    A per-item row records ONE pair, such as metres to kilograms. The units
    actually being converted may sit either side of it - a BOM in millimetres
    against stock in kilograms needs mm -> m before the factor applies, and the
    factor lands in the row's stock unit rather than whatever was asked for. So
    the chain is up to three hops:

        from -> row's alt unit -> row's stock unit -> to
        (dimensional)   (per-item factor)   (dimensional)

    Missing that middle rule is how '2 Box' became '10,000 kg' of a part held
    in grams during testing. The whole chain is multiplied out as exact
    fractions, so no rounding creeps in between the hops.

    Returns (value, status):
      'same'         - the units match, nothing to do
      'converted'    - converted through the units table
      'item_factor'  - converted using this item's own factor
      'no_unit'      - one side has no unit set
      'incompatible' - no shared family, and no factor recorded for this item

    On anything but a successful status the value is None. Callers must decide
    what to do rather than falling back to the raw number, which would silently
    compare kilograms against pieces.
    """
    if qty is None or from_unit_id is None or to_unit_id is None:
        return None, 'no_unit'
    a, b = units.get(int(from_unit_id)), units.get(int(to_unit_id))
    if not a or not b:
        return None, 'no_unit'

    q = Fraction(str(float(qty)))

    direct = _dim_factor(units, a['unit_id'], b['unit_id'])
    if direct is not None:
        status = 'same' if a['unit_id'] == b['unit_id'] else 'converted'
        return float(q * direct), status

    if not conversions or item_id is None:
        return None, 'incompatible'
    item_id = int(item_id)

    for (row_item, row_unit), row in conversions.items():
        if row_item != item_id:
            continue
        # A zero or negative factor would collapse every quantity to nothing.
        # Refuse rather than convert; a missing conversion is recoverable, a
        # silently wrong one is not.
        if row['num'] <= 0 or row['denom'] <= 0:
            continue
        factor = Fraction(row['num'], row['denom'])

        # Forward: into the row's alt unit, apply the factor, out to the target.
        f_in  = _dim_factor(units, a['unit_id'], row_unit)
        f_out = _dim_factor(units, row['stock_unit_id'], b['unit_id'])
        if f_in is not None and f_out is not None:
            return float(q * f_in * factor * f_out), 'item_factor'

        # Reverse: into the row's stock unit, divide the factor back out.
        g_in  = _dim_factor(units, a['unit_id'], row['stock_unit_id'])
        g_out = _dim_factor(units, row_unit, b['unit_id'])
        if g_in is not None and g_out is not None:
            return float(q * g_in / factor * g_out), 'item_factor'

    return None, 'incompatible'


# ─── BOM QUANTITY UNIT ──────────────────────────────────
# A BOM line does not record what unit its quantity is in, and we are not
# adding a column for it. So the rule is a convention, stated here once:
#
#   A BOM quantity is in the ITEM'S OWN UNIT, except for items bought by
#   weight and consumed by length — copper wire — where it is millimetres.
#
# Confirmed with the store 29 Sep:
#   cable, held in metres  -> BOM quantity is METRES. 0.25 means 25 cm.
#   copper wire, in grams  -> BOM quantity is MILLIMETRES, reaching grams
#                             through that item's own conversion factor.
#
# This has now been wrong in both directions, so both are worth recording.
# First it keyed off 'has a factor', which was right but looked accidental.
# Then it keyed off 'is a length', which pulled cable in and read 0.25 m as
# 0.25 mm — a thousand times too little. It is back to the factor, which is
# what actually distinguishes the two cases: an item whose stock unit cannot
# measure a BOM line is exactly the item whose BOM line is in millimetres.
#
# The weakness is worth naming plainly. This is an assumption in code, not a
# fact recorded against each line, and it holds only while copper wire is the
# only item bought by weight and used by length. The day a second such item
# arrives whose BOM is written in metres, no code change can tell them apart
# — the line has to say what it means. A `unit_id` column on bom_item does
# that, and the code below already prefers it whenever it exists.
BOM_QTY_LENGTH_UNIT = 'mm'

def unit_id_by_short(units, short):
    """unit_id for a short name, or None if that unit does not exist."""
    if not short:
        return None
    target = str(short).strip().lower()
    for uid, u in units.items():
        if (u['short'] or '').strip().lower() == target:
            return uid
    return None


def bom_unit_for(units, conversions, item_id, explicit_unit_id=None, item_unit_id=None):
    """Which unit a BOM line's quantity is in.

    An explicit value on the row wins, if that column is ever added. Otherwise
    the convention above applies. None means the item's own unit, and needs no
    conversion at all — which is the answer for almost everything, cable and
    sleeve included.
    """
    if explicit_unit_id:
        return explicit_unit_id
    if item_id is None or not conversions:
        return None

    mm_id = unit_id_by_short(units, BOM_QTY_LENGTH_UNIT)
    if not mm_id:
        return None
    length_family = unit_base_id(units, mm_id)

    # Only an item carrying a per-item factor against a length unit — copper
    # wire. Cable is held in metres and measures its own BOM line perfectly
    # well, so it never reaches here.
    iid = int(item_id)
    for (row_item, row_unit) in conversions:
        if row_item == iid and unit_base_id(units, row_unit) == length_family:
            return mm_id

    return None


_HAS_UNIT_COL = {}


def has_unit_column(cursor, table):
    """Whether `table` has the unit column this build expects.

    Checked once per process. Lets the app run before the migration has been
    applied: without the column every BOM line is simply read as being in the
    item's own unit, which is how it behaved before.
    """
    if table in _HAS_UNIT_COL:
        return _HAS_UNIT_COL[table]
    col = 'bom_unit_id' if table == 'work_order_item' else 'unit_id'
    try:
        cursor.execute(f"SHOW COLUMNS FROM `{table}` LIKE '{col}'")
        _HAS_UNIT_COL[table] = cursor.fetchone() is not None
    except Exception:
        _HAS_UNIT_COL[table] = False
    return _HAS_UNIT_COL[table]


def bom_qty_to_stock(units, conversions, item_id, bom_qty, bom_unit_id, item_unit_id):
    """Turns a BOM quantity into the unit the item is actually stocked in.

    A BOM line may state its own unit — copper wire is specified in millimetres
    while the store holds grams. bom_unit_id NULL means the line is already in
    the item's own unit, which is the case for every ordinary part.

    Returns (value, status). On a status other than success the value is None
    and the caller must leave the requirement alone rather than guess: a
    requirement of 2.4 when it should be 0.0922 is worse than no requirement,
    because it looks reasonable.
    """
    if bom_qty is None:
        return None, 'no_unit'
    if not bom_unit_id or not item_unit_id or int(bom_unit_id) == int(item_unit_id):
        return float(bom_qty), 'same'
    return convert_qty(units, bom_qty, bom_unit_id, item_unit_id, conversions, item_id)


def _trim_number(v):
    """1000.0 -> '1000', 2.5 -> '2.5'. Keeps messages readable."""
    if v is None:
        return ''
    r = round(float(v), 6)
    return str(int(r)) if r == int(r) else ('%g' % r)


_VENDOR_COLS_CACHE = None


def vendor_columns(cursor):
    """The vendor columns this database actually has, read once."""
    global _VENDOR_COLS_CACHE
    if _VENDOR_COLS_CACHE is None:
        cursor.execute("SHOW COLUMNS FROM vendor")
        _VENDOR_COLS_CACHE = {r['Field'] for r in cursor.fetchall()}
    return _VENDOR_COLS_CACHE


def vendor_select(cursor, alias='v'):
    """SELECT fragment for the vendor fields the PO screens want.

    The vendor table has been through two shapes: one with four free-text
    address lines and an email, one with a single line plus city, pincode and
    country. Naming a column that is not there fails the whole query, which is
    how the Review, Approve and Generate PO screens all became an Internal
    Server Error at once — none of them touch vendor address data except
    through this one SELECT.

    So the list is built from the live schema. Whichever shape the table is
    in, the query runs.
    """
    have = vendor_columns(cursor)
    # tax_mode decides WHICH tax lines a printed PO carries, so it has to
    # travel with the vendor to the PDF. Same schema-aware treatment: a
    # database without the column simply does not get it.
    wanted = ['gst_no', 'ph_no', 'email_id', 'tax_mode', 'currency_id',
              'address_line_1', 'address_line_2', 'address_line_3',
              'address_line_4', 'city', 'pincode', 'country']
    return ''.join(', %s.%s' % (alias, c) for c in wanted if c in have)


def fill_vendor_fields(row):
    """Give the template every field it expects, present or not.

    A missing column becomes an empty string rather than a KeyError, and where
    the table keeps city, pincode and country as separate columns they are
    folded into the second address line — the same information, in the shape
    the PO layout asks for.
    """
    if not row:
        return row
    if not row.get('address_line_2'):
        parts = [str(row.get(k) or '').strip()
                 for k in ('city', 'pincode', 'country')]
        parts = [x for x in parts if x]
        if parts:
            row['address_line_2'] = ', '.join(parts)
    for f in ('email_id', 'address_line_1', 'address_line_2',
              'address_line_3', 'address_line_4', 'gst_no', 'ph_no'):
        if row.get(f) is None:
            row[f] = ''
    return row


def known_store_locations(cursor):
    """Every location already in use, most-used first.

    Offered as suggestions rather than a closed list: a new rack exists
    before anybody adds it to a lookup table, and refusing the receipt
    until then just moves the stock somewhere untracked.
    """
    cursor.execute("""SELECT store_location AS loc, COUNT(*) AS n
                        FROM storage
                       WHERE store_location IS NOT NULL
                         AND TRIM(store_location) <> ''
                       GROUP BY store_location
                       ORDER BY n DESC, store_location""")
    return [r['loc'] for r in cursor.fetchall()]


def set_grn_item_location(cursor, grn_item_id, location):
    """Record where a received line is being put.

    Does nothing when the column is absent or no location was given, so a
    form that does not offer the field behaves exactly as before.
    """
    loc = ' '.join(str(location or '').split())
    if not loc or not grn_item_id:
        return
    try:
        cursor.execute("UPDATE grn_item SET location = %s WHERE grn_item_id = %s",
                       (loc[:100], grn_item_id))
    except Exception:
        pass          # column not present on this schema


def stamp_grn_item_unit(cursor, grn_item_id, item_id):
    """Records the unit a receipt was made in, on the GRN line itself.

    A GRN quantity is entered in the unit the item is stocked in, which today
    is only implied. If anybody ever changes an item's unit, every historical
    receipt would silently start meaning something else — 2500 recorded as
    grams would read as kilograms. Stamping the unit at the time of receipt
    makes the record say what it meant.

    Does nothing if the column has not been added yet, so the app runs either
    side of the migration.
    """
    global _GRN_ITEM_HAS_UNIT
    if _GRN_ITEM_HAS_UNIT is None:
        cursor.execute("SHOW COLUMNS FROM grn_item LIKE 'unit_id'")
        _GRN_ITEM_HAS_UNIT = cursor.fetchone() is not None
    if not _GRN_ITEM_HAS_UNIT or not grn_item_id:
        return
    cursor.execute("""UPDATE grn_item gi
                      JOIN items i ON i.item_id = %s
                      SET gi.unit_id = i.unit_id
                      WHERE gi.grn_item_id = %s AND gi.unit_id IS NULL""",
                   (item_id, grn_item_id))


_GRN_ITEM_HAS_UNIT = None


def unit_short(units, unit_id):
    u = units.get(int(unit_id)) if unit_id else None
    return u['short'] if u else ''


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
        qty_required = round(bom_qty * float(qty_to_build), QTY_DP)
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



def apply_issue_to_stock(cursor, item_id, wo_id, qty):
    """Everything that must happen to stock when quantity is issued.

    Shared by the work order Issue button and the manual issue form so the
    two paths cannot drift apart. Two steps, in order:
      1. Release this WO's blocks oldest-first, up to the quantity issued.
         Issuing consumes reserved stock first; any remainder stays blocked.
      2. Reduce physical stock in storage by the full issued quantity.
    Storage holds one row per item (enforced by uq_storage_item at startup),
    so the single UPDATE below is correct.

    Returns True if the stock was there and has been taken, False if it was
    not. A False means the caller must roll back — the blocks released in
    step 1 have to go back too.
    """
    if wo_id:
        remaining_to_release = qty
        cursor.execute("""
            SELECT stock_block_id, qty_blocked FROM stock_blocking
            WHERE item_id=%s AND wo_id=%s AND block_status=0
            ORDER BY stock_block_id ASC
        """, (item_id, wo_id))
        for block in cursor.fetchall():
            if remaining_to_release <= 0:
                break
            block_qty = float(block['qty_blocked'])
            if block_qty <= remaining_to_release:
                cursor.execute("UPDATE stock_blocking SET block_status=1 WHERE stock_block_id=%s",
                               (block['stock_block_id'],))
                remaining_to_release -= block_qty
            else:
                new_qty = round(block_qty - remaining_to_release, QTY_DP)
                cursor.execute("UPDATE stock_blocking SET qty_blocked=%s WHERE stock_block_id=%s",
                               (new_qty, block['stock_block_id']))
                remaining_to_release = 0.0

    # ─── THE STOCK CHECK LIVES HERE, NOT IN PYTHON ──────────────
    #
    # The caller reads availability, compares it, then calls this. Between
    # those two moments another request can do the same thing, and both pass
    # a check that was true when each of them looked. Two issues of 60
    # against 100 in stock both succeed, and 120 leaves a store that held
    # 100. A double-clicked Issue button does it just as easily as two
    # people.
    #
    # GREATEST(0, ...) made that worse rather than better: it clamped the
    # result to zero, so the figure looked merely empty instead of wrong and
    # the "nothing went negative" check still passed.
    #
    # Now the condition is part of the UPDATE. InnoDB locks the row for the
    # duration, so two of these serialise: the second sees the reduced
    # figure, its WHERE fails, and it changes nothing. rowcount says which
    # happened.
    cursor.execute("""UPDATE storage
                         SET physical_availability = physical_availability - %s
                       WHERE item_id = %s
                         AND physical_availability >= %s""",
                   (qty, item_id, qty))
    return cursor.rowcount > 0


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
    available = max(0.0, round(total_in - total_blocked, QTY_DP))
    return {
        'total_in': round(total_in, QTY_DP),
        'total_out': round(total_out, QTY_DP),
        'blocked': round(total_blocked, QTY_DP),
        'available': round(available, QTY_DP)
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

def currency_choices(cursor):
    """The currencies this database knows, for the picker on a PO."""
    try:
        cursor.execute("""SELECT curr_id, curr_code, curr_name
                            FROM currency
                           WHERE COALESCE(curr_is_deleted, 0) = 0
                           ORDER BY curr_id""")
        rows = cursor.fetchall()
    except Exception:
        rows = []
    if not rows:
        # The table is missing or empty: fall back to what the form used to
        # offer, so the page still works.
        rows = [{'curr_id': 0, 'curr_code': c, 'curr_name': c}
                for c in ('INR', 'USD', 'EUR', 'GBP')]
    return rows


def vendors_for_picker(cursor):
    """Vendors for a PO, each carrying the currency it invoices in.

    The currency is a property of the vendor, so the order defaults to it
    rather than to INR. Written to work whether or not vendor.currency_id
    exists, because adding that column was a separate decision.
    """
    if 'currency_id' in vendor_columns(cursor):
        sql = """SELECT v.vendor_id, v.short_name, v.full_name,
                        v.currency_id, c.curr_code
                   FROM vendor v
                   LEFT JOIN currency c ON c.curr_id = v.currency_id
                  WHERE v.is_deleted = 0
                  ORDER BY v.short_name"""
    else:
        sql = """SELECT vendor_id, short_name, full_name,
                        NULL AS currency_id, NULL AS curr_code
                   FROM vendor
                  WHERE is_deleted = 0
                  ORDER BY short_name"""
    cursor.execute(sql)
    return cursor.fetchall()


def form_item_id(form):
    """The item chosen on an HSN form, or None when none was.

    An HSN code exists whether or not anything is classified under it yet —
    the code is a fact about the goods, not about one part — so the item is
    optional on both the add and the edit screen.
    """
    raw = (form.get('item_id') or '').strip()
    if not raw:
        return None
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def may_approve_po():
    """Whether this user may review, approve, cancel, short close or reopen.

    One place, so the rule is read the same way everywhere and there is a
    single thing to change when roles come back.
    """
    if not ENFORCE_ROLE_CHECKS:
        return True
    return session.get('role') in ('admin', 'approver')


def hsn_for_item(cursor, item_id):
    """The HSN linked to this item, or None.

    One per item — hsn_items carries UNIQUE KEY uq_hsn_items_item — so there
    is never a choice to make here.
    """
    cursor.execute("""SELECT h.hsn_id, h.hsn_code, h.description, h.tax_rate
                        FROM hsn_items hi
                        JOIN hsn h ON h.hsn_id = hi.hsn_id
                       WHERE hi.item_id = %s
                         AND COALESCE(h.is_deleted, 0) = 0
                       LIMIT 1""", (item_id,))
    return cursor.fetchone()


def hsn_choices(cursor):
    """Every live HSN code, for the picker on a PO line."""
    cursor.execute("""SELECT hsn_id, hsn_code, description, tax_rate
                        FROM hsn
                       WHERE COALESCE(is_deleted, 0) = 0
                       ORDER BY hsn_code""")
    rows = cursor.fetchall()
    # tojson would otherwise hand the page a DECIMAL as a quoted string.
    for r in rows:
        r['tax_rate'] = float(r['tax_rate']) if r['tax_rate'] is not None else None
    return rows


def resolve_po_line_hsn(cursor, item_ids, chosen_ids):
    """One hsn_id per line index, plus the item codes still without one.

    An item that already has an HSN keeps it. An item with none takes the
    code the buyer picked on the line, and that choice is written to
    hsn_items, so it is asked for once rather than on every order.

    Nothing is guessed. This used to fall back to the lowest hsn_id in the
    table, which meant an unclassified item printed on a tax document at
    whatever rate happened to sit at the top of the hsn list.
    """
    cursor.execute("SELECT hsn_id FROM hsn WHERE COALESCE(is_deleted, 0) = 0")
    valid = {int(r['hsn_id']) for r in cursor.fetchall()}

    resolved, missing = {}, []
    for idx, item_id in enumerate(item_ids):
        if not item_id:
            continue

        row = hsn_for_item(cursor, item_id)
        if row:
            resolved[idx] = int(row['hsn_id'])
            continue

        picked = None
        if idx < len(chosen_ids):
            try:
                picked = int(str(chosen_ids[idx]).strip())
            except (TypeError, ValueError):
                picked = None

        if picked in valid:
            cursor.execute("""INSERT INTO hsn_items (item_id, hsn_id) VALUES (%s, %s)
                              ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)""",
                           (item_id, picked))
            resolved[idx] = picked
        else:
            cursor.execute("SELECT item_code FROM items WHERE item_id = %s", (item_id,))
            r = cursor.fetchone()
            missing.append((r or {}).get('item_code') or ('item %s' % item_id))

    return resolved, missing


def set_item_hsn(cursor, item_id, hsn_id):
    """Classify an item, if it is not classified already.

    Returns the hsn_id now on the item, or None when nothing usable was
    given and the item still has none. An item that already has an HSN is
    left alone — reclassifying belongs on the HSN screen, not as a side
    effect of entering stock.
    """
    existing = hsn_for_item(cursor, item_id)
    if existing:
        return int(existing['hsn_id'])

    try:
        wanted = int(str(hsn_id).strip())
    except (TypeError, ValueError):
        return None

    cursor.execute("""SELECT hsn_id FROM hsn
                       WHERE hsn_id = %s AND COALESCE(is_deleted, 0) = 0""", (wanted,))
    if not cursor.fetchone():
        return None

    cursor.execute("""INSERT INTO hsn_items (item_id, hsn_id) VALUES (%s, %s)
                      ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)""",
                   (item_id, wanted))
    return wanted


def no_hsn_note(missing):
    """Said after the order is saved, when a line carries no HSN.

    Not an error: the order is real and placed. But the printed copy will show
    no tax against those lines, and whoever sends it should know that before
    the vendor does.
    """
    shown = ', '.join(missing[:8])
    if len(missing) > 8:
        shown += ' and %d more' % (len(missing) - 8)
    return ("Saved. No HSN code is set for %s, so the printed order shows no "
            "tax on %s. Set one on the item when you know it." %
            (shown, 'that line' if len(missing) == 1 else 'those lines'))


def missing_hsn_message(missing):
    """What to tell the buyer about the lines that cannot be priced."""
    shown = ', '.join(missing[:8])
    if len(missing) > 8:
        shown += ' and %d more' % (len(missing) - 8)
    return ("No HSN code for %s. A purchase order cannot show tax without one, "
            "so pick an HSN on the line before saving." % shown)


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
        
        # One HSN per line. An item already classified keeps its code; one
        # that is not asks the buyer, and the answer is remembered. A line
        # with neither stops the save rather than borrowing a rate.
        hsn_id_list = request.form.getlist('hsn_id[]')
        line_hsn, missing_hsn = resolve_po_line_hsn(cursor, item_ids, hsn_id_list)
        if missing_hsn and REQUIRE_HSN_ON_PO:
            db.rollback()
            cursor.close()
            db.close()
            flash(missing_hsn_message(missing_hsn), 'error')
            return redirect(request.url)

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty = float(qty_ordered_list[i] or 0)
            price = float(unit_price_list[i] or 0)
            unit_id = int(unit_id_list[i])
            
            # None when the item has no HSN and none was picked. The line is
            # still ordered — it just carries no tax classification, and the
            # printed PO says so instead of guessing a rate.
            line_hsn_id = line_hsn.get(i)

            cursor.execute("""INSERT INTO po_items 
                             (po_id, item_id, hsn_id, qty_ordered, unit_price, unit_id, status)
                             VALUES (%s, %s, %s, %s, %s, %s, 0)""",
                           (po_id, item_id, line_hsn_id, qty, price, unit_id))

        db.commit()
        cursor.close()
        db.close()
        
        formatted_po = format_po_number(int(po_number), date_raised, po_version_number)
        flash(f'Purchase Order {formatted_po} created successfully.', 'success')
        if missing_hsn:
            flash(no_hsn_note(missing_hsn), 'warning')
        return redirect(url_for('purchase_order_list'))

    # Suggest next sequence number in current financial year
    next_po = generate_po_number(cursor)
    next_formatted_po = format_po_number(next_po, date.today(), 0)

    # Get vendors
    vendors = vendors_for_picker(cursor)
    
    # Get items for dropdown
    cursor.execute("SELECT item_id, item_code, `desc`, unit_id FROM items WHERE is_deleted=0 ORDER BY item_code")
    items = cursor.fetchall()
    
    # Get units for dropdown
    cursor.execute("SELECT unit_id, unit_short_name FROM units ORDER BY unit_name")
    units = cursor.fetchall()

    vendor_items = get_vendor_items(cursor)
    hsns = hsn_choices(cursor)
    currencies = currency_choices(cursor)

    cursor.close()
    db.close()

    return render_template('purchase_order_add.html', 
                           vendors=vendors, 
                           next_po=next_po, 
                           next_formatted_po=next_formatted_po, 
                           today=date.today(),
                           items=items,
                           units=units,
                           vendor_items=vendor_items,
                           hsns=hsns,
                           currencies=currencies)

# ─── GET VENDOR PRICING FOR AUTOFILL ────────────────────
@app.route('/purchase-order/get-vendor-pricing')
def get_vendor_pricing():
    """Price and unit for an item on a new PO line.

    Only ~1% of items have a negotiated price in item_vendor_pricing, so a
    contract-price-only lookup leaves the buyer typing everything by hand.
    Falls back through what the system already knows, most specific first,
    and reports WHICH source was used so the buyer can judge the number:

      price : contract for this vendor -> last PO to this vendor
              -> last PO to any vendor -> nothing
      unit  : the vendor's selling unit (item_vendor) -> item master.
              Deliberately NOT from a past PO line — see the note further
              down, where the unit is resolved.

    vendor_id is optional: picking the item before the vendor still fills in
    what can be known without one.
    """
    vendor_id = (request.args.get('vendor_id') or '').strip()
    item_id = (request.args.get('item_id') or '').strip()
    if not item_id:
        return jsonify(success=False, error="Missing item")

    db = get_db()
    cursor = db.cursor(dictionary=True)

    price = None
    unit_id = None
    qty = None
    price_source = None
    price_detail = None

    # 1. Negotiated price for this vendor
    if vendor_id:
        cursor.execute("""
            SELECT iv.unit_id, ivp.qty, ivp.price
            FROM item_vendor iv
            LEFT JOIN item_vendor_pricing ivp
                   ON iv.item_vendor_id = ivp.item_vendor_id AND ivp.is_deleted = 0
            WHERE iv.item_id = %s AND iv.vendor_id = %s AND iv.is_deleted = 0
            ORDER BY ivp.qty ASC
            LIMIT 1
        """, (item_id, vendor_id))
        row = cursor.fetchone()
        if row:
            unit_id = row['unit_id']
            if row['qty'] is not None:
                qty = float(row['qty'])
            if row['price'] is not None:
                price = float(row['price'])
                price_source = 'contract'
                price_detail = 'agreed price for this vendor'

    # 2. Last PO raised on this vendor for this item
    if price is None and vendor_id:
        cursor.execute("""
            SELECT poi.unit_price, poi.unit_id, po.po_number, po.date_raised
            FROM po_items poi
            JOIN purchase_order po ON po.po_id = poi.po_id
            WHERE poi.item_id = %s AND po.vendor_id = %s AND poi.unit_price > 0
            ORDER BY po.date_raised DESC, po.po_id DESC
            LIMIT 1
        """, (item_id, vendor_id))
        row = cursor.fetchone()
        if row:
            price = float(row['unit_price'])
            price_source = 'last_po_vendor'
            price_detail = f"last paid to this vendor on {row['po_number']}"

    # 3. Last PO for this item, whoever supplied it
    if price is None:
        cursor.execute("""
            SELECT poi.unit_price, poi.unit_id, po.po_number, po.date_raised,
                   COALESCE(v.short_name, v.full_name) AS vendor_name
            FROM po_items poi
            JOIN purchase_order po ON po.po_id = poi.po_id
            LEFT JOIN vendor v     ON v.vendor_id = po.vendor_id
            WHERE poi.item_id = %s AND poi.unit_price > 0
            ORDER BY po.date_raised DESC, po.po_id DESC
            LIMIT 1
        """, (item_id,))
        row = cursor.fetchone()
        if row:
            price = float(row['unit_price'])
            price_source = 'last_po_any'
            price_detail = (f"last paid on {row['po_number']}"
                            + (f" to {row['vendor_name']}" if row['vendor_name'] else ''))

    # Unit precedence is deliberately NOT the same as price precedence.
    #
    # A price is negotiated, so the last price paid is the best guess. A unit is
    # a property of the part, so history is the WORST guess: if one buyer once
    # ordered a piece-part in kilograms, taking the unit from that PO copies the
    # mistake onto every future order of the same item. That is exactly how POs
    # 1008 and 1009 came to be raised in kg for an item stocked in pieces.
    #
    # So: the vendor's selling unit if we hold one, otherwise the item master,
    # and never a past purchase order.
    if unit_id is None:
        cursor.execute("SELECT unit_id FROM items WHERE item_id = %s", (item_id,))
        row = cursor.fetchone()
        if row and row['unit_id']:
            unit_id = row['unit_id']

    # What the item is stocked in, so the form can show the conversion.
    cursor.execute("SELECT unit_id FROM items WHERE item_id = %s", (item_id,))
    row = cursor.fetchone()
    item_unit_id = row['unit_id'] if row and row['unit_id'] else None

    # The form shows what the chosen unit means in stock terms, so a buyer who
    # picks a different unit sees the consequence instead of discovering it at
    # the GRN. Nothing here restricts the choice.
    units = load_unit_map(cursor)
    conversions = load_item_conversions(cursor, [item_id])

    # When the buyer has picked a unit themselves, the note must describe THAT
    # unit, not the one we would have suggested.
    chosen = request.args.get('chosen_unit_id')
    if chosen and str(chosen).isdigit() and int(chosen) in units:
        unit_id = int(chosen)

    conv_note = None
    if unit_id and item_unit_id:
        if int(unit_id) != int(item_unit_id):
            one, status = convert_qty(units, 1, unit_id, item_unit_id,
                                      conversions, item_id)
            if status in ('converted', 'item_factor'):
                conv_note = (f"1 {unit_short(units, unit_id)} = "
                             f"{_trim_number(one)} {unit_short(units, item_unit_id)} in stock")
            elif status == 'incompatible':
                conv_note = (f"{unit_short(units, unit_id)} and "
                             f"{unit_short(units, item_unit_id)} are not the same kind of "
                             f"measure — the receipt limit cannot be enforced on this line")
    elif not item_unit_id:
        conv_note = "This item has no stock unit set, so receipts cannot be checked against the order"

    # The PO line needs an HSN. If the item has none the page asks for one
    # rather than letting the server pick, so it has to know which it is.
    hsn = hsn_for_item(cursor, item_id)

    cursor.close()
    db.close()

    return jsonify(
        success=True,
        hsn_id=(hsn['hsn_id'] if hsn else None),
        hsn_code=(hsn['hsn_code'] if hsn else None),
        hsn_rate=(float(hsn['tax_rate']) if hsn and hsn['tax_rate'] is not None else None),
        has_price_or_unit=(price is not None or unit_id is not None),
        unit_id=unit_id,
        item_unit_id=item_unit_id,
        item_unit=unit_short(units, item_unit_id) if item_unit_id else None,
        conv_note=conv_note,
        qty=qty,
        price=price,
        price_source=price_source,
        price_detail=price_detail
    )


# ─── PURCHASE ORDER DETAIL ──────────────────────────────
@app.route('/purchase-order/<int:po_id>')
def purchase_order_detail(po_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("""SELECT po.*, v.short_name as vendor_name,
                             v.full_name as vendor_full_name""" + vendor_select(cursor) + """
                      FROM purchase_order po
                      JOIN vendor v ON po.vendor_id = v.vendor_id
                      WHERE po.po_id = %s""", (po_id,))
    po = fill_vendor_fields(cursor.fetchone())
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
                             i.unit_id AS item_unit_id, iu.unit_short_name AS item_unit,
                             h.hsn_code, h.tax_rate, h.cgst, h.sgst, h.igst
                      FROM po_items pi
                      JOIN items i ON pi.item_id = i.item_id
                      LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                      LEFT JOIN units u  ON pi.unit_id = u.unit_id
                      LEFT JOIN units iu ON i.unit_id  = iu.unit_id
                      LEFT JOIN hsn h ON pi.hsn_id = h.hsn_id
                      WHERE pi.po_id = %s""", (po_id,))
    po_items = cursor.fetchall()

    # What actually came in against this PO, rolled up across every GRN.
    cursor.execute("""
        SELECT gi.item_id,
               SUM(gi.qty_received) AS received,
               SUM(gi.qty_accepted) AS accepted,
               SUM(gi.qty_rejected) AS rejected
        FROM grn_item gi
        JOIN grn g ON g.grn_id = gi.grn_id
        WHERE g.po_id = %s
        GROUP BY gi.item_id
    """, (po_id,))
    received_by_item = {int(r['item_id']): r for r in cursor.fetchall()}

    units = load_unit_map(cursor)
    conversions = load_item_conversions(cursor, [it['item_id'] for it in po_items])
    totals = {'ordered': 0.0, 'received': 0.0, 'accepted': 0.0,
              'rejected': 0.0, 'awaiting': 0.0, 'forgone': 0.0}
    for it in po_items:
        r = received_by_item.get(int(it['item_id']))
        ordered = float(it.get('qty_ordered') or 0)
        received = float(r['received'] or 0) if r else 0.0
        accepted = float(r['accepted'] or 0) if r else 0.0
        rejected = float(r['rejected'] or 0) if r else 0.0
        awaiting = max(0.0, received - accepted - rejected)

        # Received, accepted and rejected all come from GRNs and are in the
        # item's stock unit; qty_ordered is in the order's unit. Convert the
        # ordered figure across before working out what is being given up, or
        # a short close on a PO raised in kg would report a nonsense shortfall.
        ordered_conv, status = convert_qty(units, ordered,
                                           it.get('unit_id'), it.get('item_unit_id'),
                                           conversions, it.get('item_id'))
        forgone = (round(max(0.0, ordered_conv - accepted - awaiting), QTY_DP)
                   if status in ('same', 'converted', 'item_factor') else None)

        it['qty_received_total'] = round(received, QTY_DP)
        it['qty_accepted_total'] = round(accepted, QTY_DP)
        it['qty_rejected_total'] = round(rejected, QTY_DP)
        it['qty_awaiting_qc'] = round(awaiting, QTY_DP)
        it['qty_forgone'] = forgone
        it['stock_unit'] = it.get('item_unit') or ''
        it['units_comparable'] = status in ('same', 'converted', 'item_factor')
        totals['ordered'] += ordered
        totals['received'] += received
        totals['accepted'] += accepted
        totals['rejected'] += rejected
        totals['awaiting'] += awaiting
        totals['forgone'] += (forgone or 0.0)
    totals = {k: round(v, QTY_DP) for k, v in totals.items()}

    cursor.close()
    db.close()
    return render_template('purchase_order_detail.html', po=po, po_items=po_items,
                           totals=totals)

# ─── PURCHASE ORDER ACTION: GENERATE PO PDF ─────────────
@app.route('/purchase-order/<int:po_id>/po-pdf')
def purchase_order_po_pdf(po_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("""SELECT po.*, v.short_name, v.full_name""" + vendor_select(cursor) + """
                      FROM purchase_order po
                      JOIN vendor v ON po.vendor_id = v.vendor_id
                      WHERE po.po_id = %s""", (po_id,))
    po = fill_vendor_fields(cursor.fetchone())
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

    # Read BEFORE the connection closes. This sat below cursor.close(), so
    # get_setting() was handed a dead cursor, swallowed the error in its own
    # except, and quietly returned the default — meaning the policy line on
    # every printed PO was the hardcoded fallback, and changing the setting
    # in the application had no effect that anybody could see.
    policy_line_text = get_setting('policy_line_text',
                                   'COVERED UNDER NEW INDIA ASSURANCE POLICY NO. '
                                   '67020021200200000030', cursor=cursor)

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
    if po.get('po_status') not in ('Approved', SHORT_CLOSED_STATUS) and mode != 'approve':
        hide_signature = True

    # Override approved_by and approved_date dynamically for approval preview if not approved in DB yet
    if mode == 'approve' and not po.get('approved_by'):
        from datetime import date
        po['approved_by'] = session.get('full_name') or session.get('username') or 'admin'
        po['approved_date'] = date.today()

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
        # 'Short Closed' reuses the cancel_* columns to record why/who/when, so
        # it must be treated like 'Cancelled' here or an edit would wipe them.
        closed_statuses = ('Cancelled', SHORT_CLOSED_STATUS)
        if po_status in closed_statuses and not cancelled_at:
            from datetime import datetime
            cancelled_at = datetime.now()
        elif po_status not in closed_statuses:
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
        
        # One HSN per line. An item already classified keeps its code; one
        # that is not asks the buyer, and the answer is remembered. A line
        # with neither stops the save rather than borrowing a rate.
        hsn_id_list = request.form.getlist('hsn_id[]')
        line_hsn, missing_hsn = resolve_po_line_hsn(cursor, item_ids, hsn_id_list)
        if missing_hsn and REQUIRE_HSN_ON_PO:
            db.rollback()
            cursor.close()
            db.close()
            flash(missing_hsn_message(missing_hsn), 'error')
            return redirect(request.url)

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty = float(qty_ordered_list[i] or 0)
            price = float(unit_price_list[i] or 0)
            unit_id = int(unit_id_list[i])
            
            # None when the item has no HSN and none was picked. The line is
            # still ordered — it just carries no tax classification, and the
            # printed PO says so instead of guessing a rate.
            line_hsn_id = line_hsn.get(i)

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
    vendors = vendors_for_picker(cursor)
    
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

    hsns = hsn_choices(cursor)
    currencies = currency_choices(cursor)

    cursor.close()
    db.close()

    return render_template('purchase_order_edit.html', po=po, vendors=vendors, items=items, units=units, po_items=po_items, is_amend=is_amend, vendor_items=vendor_items, hsns=hsns, currencies=currencies)


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
    if not may_approve_po():
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



# ─── PURCHASE ORDER ACTION: CANCEL (BEFORE APPROVAL) ────
@app.route('/purchase-order/<int:po_id>/cancel', methods=['POST'])
def purchase_order_cancel(po_id):
    """Discards a purchase order that never reached approval.

    Distinct from short close: nothing was committed to the vendor and nothing
    can have been received, so this voids the order rather than ending it
    early. Refused once a PO is approved — that case is a short close.
    """
    if not may_approve_po():
        flash("You do not have authority to cancel a Purchase Order.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    reason = (request.form.get('cancel_reason') or '').strip()
    note = (request.form.get('cancel_note') or '').strip()
    if not reason:
        flash("Please give a reason for cancelling this Purchase Order.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT po_status FROM purchase_order WHERE po_id = %s", (po_id,))
    po = cursor.fetchone()
    if not po:
        cursor.close()
        db.close()
        return "Purchase Order not found", 404

    status = po['po_status'] or 'Draft'
    if status == 'Cancelled':
        cursor.close()
        db.close()
        flash("This Purchase Order is already cancelled.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))
    if status in ('Approved', SHORT_CLOSED_STATUS):
        cursor.close()
        db.close()
        flash("An approved Purchase Order cannot be cancelled — use Short Close instead.",
              "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    who = session.get('full_name') or session.get('username') or 'admin'
    if note:
        cursor.execute("SELECT remarks FROM purchase_order WHERE po_id = %s", (po_id,))
        existing = (cursor.fetchone() or {}).get('remarks') or ''
        stamped = f"[Cancelled {date.today()}] {note}"
        new_remarks = (existing.rstrip() + "\n" + stamped) if existing.strip() else stamped
        cursor.execute("UPDATE purchase_order SET remarks = %s WHERE po_id = %s",
                       (new_remarks, po_id))

    cursor.execute("""UPDATE purchase_order
                      SET po_status = 'Cancelled', cancel_type = %s,
                          cancelled_by = %s, cancelled_at = NOW()
                      WHERE po_id = %s""", (reason[:50], who[:50], po_id))
    db.commit()
    cursor.close()
    db.close()
    flash("Purchase Order cancelled.", "success")
    return redirect(url_for('purchase_order_detail', po_id=po_id))


# ─── PURCHASE ORDER ACTION: SHORT CLOSE ─────────────────
# Uses existing columns only — po_status carries the state (it is a VARCHAR,
# not an ENUM, so no new value needs a schema change), and the cancel_* trio
# records why, who and when. Because the GRN form already filters on
# po_status = 'Approved', a short-closed PO drops out of it automatically.
@app.route('/purchase-order/<int:po_id>/short-close', methods=['POST'])
def purchase_order_short_close(po_id):
    """Ends an approved PO with less than the ordered quantity received.

    Covers both cases the business described — the vendor cannot supply the
    balance, and we have received enough — because they end the same way. The
    difference is recorded as the reason, not as a separate state. Valid with
    nothing received at all.
    """
    if not may_approve_po():
        flash("You do not have authority to short close a Purchase Order.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    reason = (request.form.get('short_close_reason') or '').strip()
    note = (request.form.get('short_close_note') or '').strip()
    if not reason:
        flash("Please give a reason for short closing this Purchase Order.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT po_status FROM purchase_order WHERE po_id = %s", (po_id,))
    po = cursor.fetchone()
    if not po:
        cursor.close()
        db.close()
        return "Purchase Order not found", 404

    if po['po_status'] == SHORT_CLOSED_STATUS:
        cursor.close()
        db.close()
        flash("This Purchase Order is already short closed.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    if po['po_status'] != 'Approved':
        cursor.close()
        db.close()
        flash("Only an approved Purchase Order can be short closed.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    who = session.get('full_name') or session.get('username') or 'admin'
    if note:
        # remarks is TEXT, so the note survives in full; appended rather than
        # replaced so anything already noted on the PO is kept.
        cursor.execute("SELECT remarks FROM purchase_order WHERE po_id = %s", (po_id,))
        existing = (cursor.fetchone() or {}).get('remarks') or ''
        stamped = f"[Short close {date.today()}] {note}"
        new_remarks = (existing.rstrip() + "\n" + stamped) if existing.strip() else stamped
        cursor.execute("UPDATE purchase_order SET remarks = %s WHERE po_id = %s",
                       (new_remarks, po_id))

    cursor.execute("""UPDATE purchase_order
                      SET po_status = %s, cancel_type = %s,
                          cancelled_by = %s, cancelled_at = NOW()
                      WHERE po_id = %s""",
                   (SHORT_CLOSED_STATUS, reason[:50], who[:50], po_id))
    db.commit()
    cursor.close()
    db.close()
    flash("Purchase Order short closed. It can no longer be selected for a new GRN.",
          "success")
    return redirect(url_for('purchase_order_detail', po_id=po_id))


# ─── PURCHASE ORDER ACTION: REOPEN AFTER SHORT CLOSE ────
@app.route('/purchase-order/<int:po_id>/reopen', methods=['POST'])
def purchase_order_reopen(po_id):
    """Puts a short-closed PO back to Approved so it can be received against
    again — the vendor can supply after all, or the requirement returned."""
    if not may_approve_po():
        flash("You do not have authority to reopen a Purchase Order.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT po_status FROM purchase_order WHERE po_id = %s", (po_id,))
    po = cursor.fetchone()
    if not po:
        cursor.close()
        db.close()
        return "Purchase Order not found", 404
    if po['po_status'] != SHORT_CLOSED_STATUS:
        cursor.close()
        db.close()
        flash("This Purchase Order is not short closed.", "error")
        return redirect(url_for('purchase_order_detail', po_id=po_id))

    cursor.execute("""UPDATE purchase_order
                      SET po_status = 'Approved', cancel_type = NULL,
                          cancelled_by = NULL, cancelled_at = NULL
                      WHERE po_id = %s""", (po_id,))
    db.commit()
    cursor.close()
    db.close()
    flash("Purchase Order reopened. It can be selected for a GRN again.", "success")
    return redirect(url_for('purchase_order_detail', po_id=po_id))


# ─── BOM (Bill of Materials) ────────────────────────────
@app.route('/bom')
def bom_list():
    search = request.args.get('search', '').strip()

    db = get_db()
    cursor = db.cursor(dictionary=True)

    query = """SELECT b.bom_id, b.bom_code, b.description,
                      COUNT(DISTINCT bi.bom_item_id) AS item_count,
                      COUNT(DISTINCT wo.wo_id) AS wo_count
               FROM bom b
               LEFT JOIN bom_item bi
                      ON bi.bom_id = b.bom_id AND COALESCE(bi.is_deleted, 0) = 0
               LEFT JOIN work_order wo ON wo.bom_id = b.bom_id
               WHERE COALESCE(b.is_deleted, 0) = 0"""
    params = []
    if search:
        query += " AND (b.bom_code LIKE %s OR b.description LIKE %s)"
        params.extend([f"%{search}%", f"%{search}%"])
    query += """ GROUP BY b.bom_id, b.bom_code, b.description
                 ORDER BY b.bom_code"""

    cursor.execute(query, tuple(params))
    boms = cursor.fetchall()

    cursor.close()
    db.close()
    return render_template('bom_list.html', boms=boms, search=search)


@app.route('/bom/<int:bom_id>')
def bom_detail(bom_id):
    # ?path=12,34 carries the BOMs already stepped through, so the breadcrumb
    # can show the route taken and each crumb can link back to its own level.
    raw_path = request.args.get('path', '')
    ancestors = [int(p) for p in raw_path.split(',') if p.strip().isdigit()][:25]
    # A BOM that ends up inside itself would otherwise nest forever.
    is_cycle = bom_id in ancestors

    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("""SELECT * FROM bom
                      WHERE bom_id = %s AND COALESCE(is_deleted, 0) = 0""", (bom_id,))
    bom = cursor.fetchone()
    if not bom:
        cursor.close()
        db.close()
        flash('BOM not found.', 'error')
        return redirect(url_for('bom_list'))

    cursor.execute("""SELECT bi.bom_item_id, bi.qty, bi.patch, bi.reference, bi.section,
                             i.item_id, i.item_code,
                             i.desc as item_desc, i.mpn,
                             COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name,
                             u.unit_short_name as unit,
                             (SELECT b2.bom_id FROM bom b2
                               WHERE b2.bom_code = i.item_code
                                 AND COALESCE(b2.is_deleted, 0) = 0
                               LIMIT 1) AS child_bom_id
                      FROM bom_item bi
                      JOIN items i ON bi.item_id = i.item_id
                      LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                      LEFT JOIN units u ON i.unit_id = u.unit_id
                      WHERE bi.bom_id = %s AND COALESCE(bi.is_deleted, 0) = 0
                      ORDER BY i.item_code""", (bom_id,))
    bom_items_raw = cursor.fetchall()

    bom_items = []
    dns_count = 0
    sub_bom_count = 0
    for row in bom_items_raw:
        stock = get_item_stock(cursor, row['item_id'])
        row['available'] = stock['available']
        # DNS rows stay listed here so the BOM mirrors the engineering sheet;
        # they are only excluded when a work order is built from it.
        row['is_dns'] = is_dns_patch(row.get('patch'))
        if row['is_dns']:
            dns_count += 1
        # A line is a sub-assembly when its item code is also a BOM code.
        # There is no parent/child column on bom_item — the code match is the
        # only link, which is why it is looked up per row rather than assumed
        # from a code prefix (sub-BOM codes start 4002, 4005, 5007, 5020, ...).
        if row.get('child_bom_id'):
            sub_bom_count += 1
        bom_items.append(row)

    # Work orders already built from this BOM, so the user can see what a
    # re-upload would affect before they do it.
    cursor.execute("""SELECT wo_id, wo_number, qty_to_build, status
                      FROM work_order WHERE bom_id = %s
                      ORDER BY wo_id DESC""", (bom_id,))
    linked_wos = cursor.fetchall()

    # Breadcrumb entries for the BOMs already passed through, each carrying the
    # shorter path that leads back to it.
    trail = []
    if ancestors:
        placeholders = ', '.join(['%s'] * len(ancestors))
        cursor.execute(f"SELECT bom_id, bom_code FROM bom WHERE bom_id IN ({placeholders})",
                       tuple(ancestors))
        code_by_id = {r['bom_id']: r['bom_code'] for r in cursor.fetchall()}
        for idx, anc_id in enumerate(ancestors):
            trail.append({
                'bom_id': anc_id,
                'bom_code': code_by_id.get(anc_id, anc_id),
                'path': ','.join(str(x) for x in ancestors[:idx]),
            })

    # The path a child link should carry: everything so far, plus this BOM.
    child_path = ','.join(str(x) for x in (ancestors + [bom_id]))

    cursor.close()
    db.close()
    return render_template('bom_detail.html', bom=bom, bom_items=bom_items,
                           dns_count=dns_count, sub_bom_count=sub_bom_count,
                           trail=trail, child_path=child_path, is_cycle=is_cycle,
                           linked_wos=linked_wos, status_map=STATUS_MAP)


# ─── WORK ORDER LIST ───────────────────────────────────────────────────────────────────────
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

# ─── WORK ORDER ADD ─────────────────────────────────────
@app.route('/work-order/add', methods=['GET', 'POST'])
def work_order_add():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        qty_to_build = request.form['qty_to_build']
        description = request.form.get('description', '')
        # Who created this is a fact about the session, not something the
        # browser gets to state. Taking it from the form meant a posted value
        # could name anyone, and a blank field recorded nobody at all.
        created_by = (session.get('full_name') or session.get('username')
                      or 'admin')
        remarks = request.form.get('remarks', '')
        action = request.form.get('action', 'draft')
        bom_id_raw = request.form.get('bom_id', '').strip()

        if not bom_id_raw.isdigit():
            flash('Please select a BOM.', 'error')
            cursor.close()
            db.close()
            return redirect(url_for('work_order_add'))

        bom_id = int(bom_id_raw)
        cursor.execute("""SELECT bom_id, bom_code, description FROM bom
                          WHERE bom_id = %s AND COALESCE(is_deleted, 0) = 0""", (bom_id,))
        bom_row = cursor.fetchone()
        if not bom_row:
            flash('That BOM no longer exists. Please pick another.', 'error')
            cursor.close()
            db.close()
            return redirect(url_for('work_order_add'))

        if not description and bom_row['description']:
            description = bom_row['description']

        wo_number = generate_wo_number(cursor)
        cursor.execute("""INSERT INTO work_order
                         (wo_number, bom_id, qty_to_build, description, created_by, status, remarks)
                         VALUES (%s, %s, %s, %s, %s, 0, %s)""",
                       (wo_number, bom_id, qty_to_build, description, created_by, remarks))
        wo_id = cursor.lastrowid

        if action == 'calculate':
            calculate_requirements_from_bom.last_dns_excluded = []
            calculate_requirements_from_bom.last_merged = []
            imported = calculate_requirements_from_bom(cursor, wo_id, bom_id, qty_to_build)
            excluded = getattr(calculate_requirements_from_bom, 'last_dns_excluded', [])
            db.commit()
            cursor.close()
            db.close()

            msg = (f"Work order {wo_number} created from BOM {bom_row['bom_code']} "
                   f"with {imported} item(s).")
            if excluded:
                codes = ', '.join(d['item_code'] for d in excluded[:8])
                if len(excluded) > 8:
                    codes += f", +{len(excluded) - 8} more"
                # Named, not silent: a quietly dropped item is indistinguishable
                # from a missing one when someone checks the board later.
                msg += (f" {len(excluded)} Do Not Stuff item(s) were excluded "
                        f"({codes}).")

            merged = getattr(calculate_requirements_from_bom, 'last_merged', [])
            if merged:
                msg += (f" {len(merged)} item(s) appear more than once on this BOM; "
                        f"their quantities were added together.")
            flash(msg, 'success')
            return redirect(url_for('work_order_detail', wo_id=wo_id))

        db.commit()
        cursor.close()
        db.close()
        flash(f"Work order {wo_number} saved as draft.", 'success')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

    # GET — every BOM, with its stuffable item count and DNS count.
    # Counted in Python rather than SQL because whether a patch value means
    # DNS is a text rule (is_dns_patch), not something SQL should guess at.
    # LEFT JOIN so a BOM with no items still appears, with bom_item_id NULL.
    cursor.execute("""SELECT b.bom_id, b.bom_code, b.description,
                             bi.bom_item_id, bi.item_id, bi.patch
                      FROM bom b
                      LEFT JOIN bom_item bi
                             ON bi.bom_id = b.bom_id AND COALESCE(bi.is_deleted, 0) = 0
                      WHERE COALESCE(b.is_deleted, 0) = 0
                      ORDER BY b.bom_code""")

    # Tally per item, because an item can appear on a BOM many times. An item
    # is only really excluded when EVERY one of its rows is DNS — if it has
    # even one normal row it still goes on the work order (with that row's
    # quantity), so counting DNS rows would overstate what gets left out.
    boms = {}
    per_item = {}
    for r in cursor.fetchall():
        entry = boms.setdefault(r['bom_id'], {
            'bom_id': r['bom_id'],
            'bom_code': r['bom_code'],
            'description': r['description'],
            'item_count': 0,
            'dns_count': 0,
        })
        if r['bom_item_id'] is None:
            continue                      # BOM has no items at all
        key = (r['bom_id'], r['item_id'])
        tally = per_item.setdefault(key, {'rows': 0, 'dns_rows': 0})
        tally['rows'] += 1
        if is_dns_patch(r['patch']):
            tally['dns_rows'] += 1

    for (bom_id, _item_id), tally in per_item.items():
        entry = boms.get(bom_id)
        if not entry:
            continue
        if tally['dns_rows'] == tally['rows']:
            entry['dns_count'] += 1       # every row DNS -> item is excluded
        else:
            entry['item_count'] += 1      # appears on the work order

    cursor.close()
    db.close()
    return render_template('work_order_add.html',
                           boms=sorted(boms.values(), key=lambda b: b['bom_code'] or ''))


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
    # bom_unit is what the BOM line stated; unit is what the store holds. They
    # differ for copper wire — specified in mm, held in grams — so both are
    # shown rather than leaving the reader to guess which a number is in.
    bom_unit_join = ('LEFT JOIN units bu ON bu.unit_id = wi.bom_unit_id'
                     if has_unit_column(cursor, 'work_order_item') else '')
    bom_unit_col = 'bu.unit_short_name' if bom_unit_join else 'NULL'
    cursor.execute(f"""SELECT wi.*, i.item_code, i.desc as item_desc, i.mpn,
                             COALESCE(wi.manufacturer_name, m.mfr_short_name, m.mfr_full_name) as manufacturer_name,
                             u.unit_short_name as unit,
                             i.unit_id AS item_unit_id,
                             {bom_unit_col} AS bom_unit
                     FROM work_order_item wi
                     JOIN items i ON wi.item_id = i.item_id
                     LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
                     LEFT JOIN units u ON i.unit_id = u.unit_id
                     {bom_unit_join}
                     WHERE wi.wo_id = %s AND COALESCE(wi.is_removed, 0) = 0
                     ORDER BY i.item_code""", (wo_id,))
    wo_items_raw = cursor.fetchall()

    # Label each row with the unit its BOM quantity is in, so a number in the
    # BOM Qty column cannot be mistaken for the unit beside it.
    wo_units = load_unit_map(cursor)
    wo_conversions = load_item_conversions(cursor, [r['item_id'] for r in wo_items_raw])
    for item in wo_items_raw:
        bu = bom_unit_for(wo_units, wo_conversions, item['item_id'],
                          item.get('bom_unit_id'), item.get('item_unit_id'))
        item['bom_unit'] = unit_short(wo_units, bu) if bu else None

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
        # Two separate questions, two separate columns:
        #   pending  = work still to do          -> required - issued
        #   shortage = stock still to be sourced -> pending - blocked(this WO) - available
        # An item can be fully in stock and still pending (nobody issued it yet),
        # and it can be pending with no shortage at all.
        item['pending'] = max(0.0, round(float(item['qty_required']) - float(item['qty_issued'] or 0), QTY_DP))
        item['shortage'] = max(0.0, round(item['pending'] - item['wo_blocked'] - item['available'], QTY_DP))
        # Issued beyond the requirement (scrap replacement, over-draw, a BOM qty
        # later reduced). Pending floors at zero, so without this the excess is invisible.
        item['additional_issued'] = max(0.0, round(float(item['qty_issued'] or 0) - float(item['qty_required']), QTY_DP))
        
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

    # The edited figure is in whatever unit the BOM line uses; the requirement
    # has to come back out in the unit the item is stocked in.
    u_col = 'wi.bom_unit_id' if has_unit_column(cursor, 'work_order_item') else 'NULL'
    cursor.execute(f"""SELECT wi.item_id, {u_col} AS bom_unit_id, i.unit_id AS item_unit_id
                        FROM work_order_item wi
                        JOIN items i ON i.item_id = wi.item_id
                       WHERE wi.wo_item_id = %s""", (wi_id,))
    ctx = cursor.fetchone() or {}
    units = load_unit_map(cursor)
    conversions = load_item_conversions(cursor, [ctx.get('item_id')] if ctx else None)
    bom_unit = bom_unit_for(units, conversions, ctx.get('item_id'),
                            ctx.get('bom_unit_id'), ctx.get('item_unit_id'))
    conv, status = bom_qty_to_stock(units, conversions, ctx.get('item_id'),
                                    bom_qty * float(wo['qty_to_build']),
                                    bom_unit, ctx.get('item_unit_id'))
    qty_required = round(conv, QTY_DP) if status in ('same', 'converted', 'item_factor') else None

    cursor.execute("""UPDATE work_order_item 
                     SET bom_qty=%s, qty_required=%s 
                     WHERE wo_item_id=%s AND wo_id=%s""",
                   (bom_qty, qty_required, wi_id, wo_id))

    cursor.execute("SELECT item_id, qty_issued FROM work_order_item WHERE wo_item_id=%s", (wi_id,))
    wi_row = cursor.fetchone()
    item_id = wi_row['item_id'] if wi_row else None
    qty_issued = float(wi_row['qty_issued'] or 0) if wi_row else 0.0

    cursor.execute("""
        SELECT COALESCE(SUM(qty_blocked), 0) as wo_blocked
        FROM stock_blocking
        WHERE item_id = %s
        AND wo_id=%s AND block_status=0
    """, (item_id, wo_id))
    wo_blocked = float(cursor.fetchone()['wo_blocked'] or 0)

    # Same two rules as the page render, so the live update and a refresh agree.
    pending = max(0.0, round(qty_required - qty_issued, QTY_DP))
    stock = get_item_stock(cursor, item_id) if item_id else {'available': 0.0}
    shortage = max(0.0, round(pending - wo_blocked - stock['available'], QTY_DP))
    additional_issued = max(0.0, round(qty_issued - qty_required, QTY_DP))

    db.commit()
    cursor.close()
    db.close()

    # Already-issued quantity needs no further blocking, so compare the
    # blocked amount against what is still pending.
    needs_block_alert = wo_blocked < pending
    block_diff = round(pending - wo_blocked, QTY_DP)

    return jsonify({
        'success': True,
        'qty_required': qty_required,
        'wo_blocked': wo_blocked,
        'needs_block_alert': needs_block_alert,
        'block_diff': block_diff if needs_block_alert else 0,
        'pending': pending,
        'shortage': shortage,
        'additional_issued': additional_issued
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
    qty_required = round(bom_qty * float(wo['qty_to_build']), QTY_DP)

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
            new_qty = round(block_qty - remaining, QTY_DP)
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

    if not apply_issue_to_stock(cursor, item_id, wo_id, qty_to_issue):
        # The stock was there when we looked and gone when we reached for it.
        # Somebody else took it in between — or this form was submitted twice.
        db.rollback()
        cursor.close()
        db.close()
        flash(f"Could not issue {_trim_number(qty_to_issue)} — the stock was "
              f"taken between checking and issuing. Nothing was changed. "
              f"Reload the work order to see what is actually available.",
              'error')
        return redirect(url_for('work_order_detail', wo_id=wo_id))

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
                
        elif grn_type == 1:  # Back To Store GRN
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
            # Internal Return, Online Purchase, Miscellaneous, Other
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
        invoice_no = request.form.get('invoice_no', '').strip()
        invoice_date = request.form.get('invoice_date') or str(date.today())
        remarks = request.form.get('remarks', '')
        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        # Invoice number is mandatory for Purchase GRNs only. The template also
        # marks the field required, but that is a browser-side check a user can
        # bypass, so the rule is enforced here as well.
        if not invoice_no:
            cursor.close()
            db.close()
            flash('Invoice number is required for a Purchase GRN.', 'error')
            return redirect(url_for('grn_add_purchase', grn_no=grn_no))

        item_ids = request.form.getlist('item_id[]')
        qty_received = request.form.getlist('qty_received[]')
        line_remarks = request.form.getlist('line_remarks[]')
        line_locations = request.form.getlist('location[]')

        # Cannot receive more than the PO still has outstanding. The form also
        # caps each box, but that is a browser-side check a user can bypass, so
        # the rule is enforced here — and checked BEFORE the GRN header is
        # written, so a rejected GRN leaves nothing behind.
        unchecked_lines = []
        if po_id:
            cursor.execute("""
                SELECT poi.item_id, poi.qty_ordered, poi.unit_id AS po_unit_id,
                       i.item_code, i.unit_id AS item_unit_id,
                       COALESCE(r.received, 0) AS received,
                       COALESCE(r.accepted, 0) AS accepted,
                       COALESCE(r.rejected, 0) AS rejected
                FROM po_items poi
                JOIN items i ON i.item_id = poi.item_id
                LEFT JOIN (
                    SELECT gi.item_id,
                           SUM(gi.qty_received) AS received,
                           SUM(gi.qty_accepted) AS accepted,
                           SUM(gi.qty_rejected) AS rejected
                    FROM grn_item gi
                    JOIN grn g ON g.grn_id = gi.grn_id
                    WHERE g.po_id = %s
                    GROUP BY gi.item_id
                ) r ON r.item_id = poi.item_id
                WHERE poi.po_id = %s
            """, (po_id, po_id))

            # Read this result set BEFORE anything else touches the cursor.
            # load_unit_map runs its own query, and mysql-connector raises
            # "Unread result found" if a second statement starts while rows
            # from the first are still waiting.
            po_rows = cursor.fetchall()
            units = load_unit_map(cursor)
            conversions = load_item_conversions(cursor, [r['item_id'] for r in po_rows])
            pending_by_item = {}
            for r in po_rows:
                ordered_q = float(r['qty_ordered'] or 0)
                accepted_q = float(r['accepted'] or 0)
                awaiting_q = max(0.0, float(r['received'] or 0) - accepted_q - float(r['rejected'] or 0))

                # The ordered quantity is in the ORDER's unit; accepted and
                # awaiting come from GRNs and are in the ITEM's unit. Convert
                # first, subtract second — the other way round subtracts grams
                # from kilograms.
                ordered_conv, status = convert_qty(units, ordered_q,
                                                   r['po_unit_id'], r['item_unit_id'],
                                                   conversions, r['item_id'])
                allowed = (round(max(0.0, ordered_conv - accepted_q - awaiting_q), QTY_DP)
                           if status in ('same', 'converted', 'item_factor') else None)
                pending_by_item[int(r['item_id'])] = {
                    'allowed': allowed,
                    'status': status,
                    'item_code': r['item_code'],
                    'po_unit': unit_short(units, r['po_unit_id']),
                    'item_unit': unit_short(units, r['item_unit_id']),
                }

            over = []
            for i, item_id in enumerate(item_ids):
                if not item_id:
                    continue
                try:
                    qty_r = float(qty_received[i] or 0)
                except (ValueError, IndexError):
                    qty_r = 0.0
                info = pending_by_item.get(int(item_id))
                if info is None:
                    continue          # not a line on this PO; nothing to cap against

                if info['allowed'] is None:
                    # Units that do not share a family have no factor between
                    # them, so there is no honest comparison to make. Agreed
                    # behaviour is to let the receipt through and say so rather
                    # than compare raw numbers, which would be meaningless.
                    if qty_r > 0:
                        unchecked_lines.append(
                            f"{info['item_code']} (ordered in {info['po_unit'] or '?'}, "
                            f"stocked in {info['item_unit'] or '?'})")
                    continue

                if round(qty_r, QTY_DP) > info['allowed'] + 0.0005:
                    over.append(f"{info['item_code']} (entered {_trim_number(qty_r)} "
                                f"{info['item_unit']}, outstanding "
                                f"{_trim_number(info['allowed'])} {info['item_unit']})")

            if over:
                cursor.close()
                db.close()
                flash('Cannot receive more than the purchase order has outstanding: '
                      + '; '.join(over) + '.', 'error')
                return redirect(url_for('grn_add_purchase', grn_no=grn_no))

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
            loc = line_locations[i] if i < len(line_locations) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            # Where it is being put. Recorded on the line now; storage only
            # moves once QC accepts, because rejected goods are never put away.
            set_grn_item_location(cursor, new_line_id, loc)

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
    locations = known_store_locations(cursor)
    
    cursor.close()
    db.close()
    return render_template('grn_add_purchase.html', grn_no=grn_no, purchase_orders=purchase_orders, today=date.today(), locations=locations)


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
        line_locations = request.form.getlist('location[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            # Where this line is being put away. Optional on every form:
            # a receipt with no location is still a receipt, and refusing
            # it would only push the stock somewhere untracked.
            set_grn_item_location(cursor, new_line_id,
                                  line_locations[i] if i < len(line_locations) else '')

        db.commit()
        cursor.close()
        db.close()
        flash(f'Internal Return GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    returned_by = session.get('username') or session.get('full_name') or 'admin'
    
    cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
    items = cursor.fetchall()
    locations = known_store_locations(cursor)

    cursor.close()
    db.close()
    return render_template('grn_add_internal_return.html', grn_no=grn_no, returned_by=returned_by, items=items, today=date.today(), locations=locations)


# ─── GRN ADD WORK ORDER RETURN (Step 2 for WO Return) ──────────
@app.route('/grn/add/wo-return', methods=['GET', 'POST'])
def grn_add_wo_return():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        wo_id_ref = request.form.get('wo_id_ref') or None
        invoice_no = request.form.get('invoice_no', '').strip() or None  # optional for this GRN type
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
        line_locations = request.form.getlist('location[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            # Where this line is being put away. Optional on every form:
            # a receipt with no location is still a receipt, and refusing
            # it would only push the stock somewhere untracked.
            set_grn_item_location(cursor, new_line_id,
                                  line_locations[i] if i < len(line_locations) else '')

        db.commit()
        cursor.close()
        db.close()
        flash(f'Back To Store GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("""
        SELECT wo_id, wo_number, description FROM work_order ORDER BY wo_id DESC
    """)
    work_orders = cursor.fetchall()
    locations = known_store_locations(cursor)

    cursor.close()
    db.close()
    return render_template('grn_add_wo_return.html', grn_no=grn_no, work_orders=work_orders, today=date.today(), locations=locations)


# ─── GRN ADD TYPE 5: WORK ORDER RETURN ─────────────────────────
@app.route('/grn/add/work-order-return', methods=['GET', 'POST'])
def grn_add_work_order_return():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        wo_id_ref = request.form.get('wo_id_ref') or None
        invoice_no = request.form.get('invoice_no', '').strip() or None  # optional for this GRN type
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
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            set_grn_item_location(cursor, new_line_id,
                                  request.form.get('location', ''))

        db.commit()
        cursor.close()
        db.close()
        flash(f'Work Order Output GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("""
        SELECT wo.wo_id, wo.wo_number, wo.description AS wo_description, b.bom_code, b.description AS bom_desc
        FROM work_order wo
        LEFT JOIN bom b ON wo.bom_id = b.bom_id
        ORDER BY wo.wo_id DESC
    """)
    work_orders = cursor.fetchall()
    locations = known_store_locations(cursor)

    cursor.close()
    db.close()
    return render_template('grn_add_work_order_return.html', grn_no=grn_no, work_orders=work_orders, today=date.today(), locations=locations)


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
        invoice_no = request.form.get('invoice_no', '').strip()
        invoice_date = request.form.get('invoice_date') or str(date.today())
        remarks = request.form.get('remarks', '')
        received_by = session.get('username') or session.get('full_name') or 'admin'
        received_date = invoice_date

        # Invoice number is mandatory for Online Purchase GRNs only.
        if not invoice_no:
            cursor.close()
            db.close()
            flash('Invoice number is required for an Online Purchase GRN.', 'error')
            return redirect(url_for('grn_add_online_purchase', grn_no=grn_no))

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
        line_locations = request.form.getlist('location[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            # Where this line is being put away. Optional on every form:
            # a receipt with no location is still a receipt, and refusing
            # it would only push the stock somewhere untracked.
            set_grn_item_location(cursor, new_line_id,
                                  line_locations[i] if i < len(line_locations) else '')

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
    locations = known_store_locations(cursor)

    cursor.close()
    db.close()
    return render_template('grn_add_online_purchase.html', grn_no=grn_no, vendors=vendors, items=items, purchase_orders=purchase_orders, today=date.today(), locations=locations)


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


# ─── GRN ADD MISCELLANEOUS (Step 2 for Miscellaneous, grn_type = 4) ───────
@app.route('/grn/add/any-type', methods=['GET', 'POST'])
def grn_add_any_type():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    if request.method == 'POST':
        grn_no = request.form.get('grn_no') or generate_grn_number(cursor)
        v_input = (request.form.get('vendor_id') or '').strip()
        vendor_id = int(v_input) if v_input and v_input.isdigit() and int(v_input) > 0 else None
        invoice_no = request.form.get('invoice_no', '').strip() or None  # optional for this GRN type
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
        line_locations = request.form.getlist('location[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            # Where this line is being put away. Optional on every form:
            # a receipt with no location is still a receipt, and refusing
            # it would only push the stock somewhere untracked.
            set_grn_item_location(cursor, new_line_id,
                                  line_locations[i] if i < len(line_locations) else '')

        db.commit()
        cursor.close()
        db.close()
        flash(f'Miscellaneous GRN {grn_no} created successfully.', 'success')
        return redirect(url_for('grn_list'))

    grn_no = request.args.get('grn_no') or generate_grn_number(cursor)
    
    cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted = 0 ORDER BY short_name")
    vendors = cursor.fetchall()

    cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
    items = cursor.fetchall()
    locations = known_store_locations(cursor)

    cursor.close()
    db.close()
    return render_template('grn_add_any_type.html', grn_no=grn_no, vendors=vendors, items=items, today=date.today(), locations=locations)


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
        line_locations = request.form.getlist('location[]')

        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            # Where this line is being put away. Optional on every form:
            # a receipt with no location is still a receipt, and refusing
            # it would only push the stock somewhere untracked.
            set_grn_item_location(cursor, new_line_id,
                                  line_locations[i] if i < len(line_locations) else '')

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
        locations = known_store_locations(cursor)
        cursor.close()
        db.close()
        return render_template('grn_add_purchase.html', grn_no=grn['grn_no'], purchase_orders=purchase_orders, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type, locations=locations)

    elif target_type == 1: # Back To Store
        cursor.execute("SELECT wo_id, wo_number, description FROM work_order ORDER BY wo_id DESC")
        work_orders = cursor.fetchall()
        locations = known_store_locations(cursor)
        cursor.close()
        db.close()
        return render_template('grn_add_wo_return.html', grn_no=grn['grn_no'], work_orders=work_orders, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type, locations=locations)

    elif target_type == 2: # Internal Return
        cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
        items = cursor.fetchall()
        returned_by = grn.get('received_by') or session.get('username') or 'admin'
        locations = known_store_locations(cursor)
        cursor.close()
        db.close()
        return render_template('grn_add_internal_return.html', grn_no=grn['grn_no'], returned_by=returned_by, items=items, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type, locations=locations)

    elif target_type == 3: # Online Purchase
        cursor.execute("SELECT vendor_id, short_name, full_name FROM vendor WHERE is_deleted = 0 ORDER BY short_name")
        vendors = cursor.fetchall()
        cursor.execute("SELECT item_id, item_code, `desc` AS item_desc, mpn FROM items WHERE is_deleted = 0 ORDER BY item_code")
        items = cursor.fetchall()
        locations = known_store_locations(cursor)
        cursor.close()
        db.close()
        return render_template('grn_add_online_purchase.html', grn_no=grn['grn_no'], vendors=vendors, items=items, today=grn['invoice_date'] or date.today(), is_convert=True, grn=grn, existing_items=existing_items, target_type=target_type, locations=locations)

    elif target_type == 5: # Work Order Output
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
        invoice_no = request.form.get('invoice_no', '').strip() or None  # optional for this GRN type
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
        line_locations = request.form.getlist('location[]')
        
        for i, item_id in enumerate(item_ids):
            if not item_id:
                continue
            qty_r = float(qty_received[i] or 0)
            rem = line_remarks[i] if i < len(line_remarks) else ''
            cursor.execute("""
                INSERT INTO grn_item (grn_id, item_id, qty_received, qty_rejected, qty_accepted, posted_qty_accepted, qc_status, remarks, created_at)
                VALUES (%s, %s, %s, 0, 0, 0, 'Pending', %s, NOW())
            """, (grn_id, item_id, qty_r, rem))
            new_line_id = cursor.lastrowid
            stamp_grn_item_unit(cursor, new_line_id, item_id)
            # Where this line is being put away. Optional on every form:
            # a receipt with no location is still a receipt, and refusing
            # it would only push the stock somewhere untracked.
            set_grn_item_location(cursor, new_line_id,
                                  line_locations[i] if i < len(line_locations) else '')

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
    locations = known_store_locations(cursor)
    
    cursor.close()
    db.close()
    return render_template('grn_add_other.html', grn_no=grn_no, grn_type=grn_type_int, grn_type_label=grn_type_label, open_wos=open_wos, items=items, vendors=vendors, today=date.today(), locations=locations)


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
    
    # Rolled up across EVERY GRN against this PO, because a line can be received
    # in several deliveries and inspected in several passes.
    cursor.execute("""
        SELECT poi.*, i.item_code, i.`desc` AS item_desc, u.unit_short_name AS unit,
               i.unit_id AS item_unit_id, iu.unit_short_name AS item_unit,
               COALESCE(r.received, 0) AS qty_received_total,
               COALESCE(r.accepted, 0) AS qty_accepted_total,
               COALESCE(r.rejected, 0) AS qty_rejected_total
        FROM po_items poi
        JOIN items i ON poi.item_id = i.item_id
        LEFT JOIN units u  ON poi.unit_id = u.unit_id
        LEFT JOIN units iu ON i.unit_id   = iu.unit_id
        LEFT JOIN (
            SELECT gi.item_id,
                   SUM(gi.qty_received) AS received,
                   SUM(gi.qty_accepted) AS accepted,
                   SUM(gi.qty_rejected) AS rejected
            FROM grn_item gi
            JOIN grn g ON g.grn_id = gi.grn_id
            WHERE g.po_id = %s
            GROUP BY gi.item_id
        ) r ON r.item_id = poi.item_id
        WHERE poi.po_id = %s
    """, (po_id, po_id))
    items = cursor.fetchall()
    # Where each item is kept today, so the receipt can default to it and
    # the storeman only types when something is actually moving.
    current_loc = {}
    ids = [it['item_id'] for it in items]
    if ids:
        ph = ','.join(['%s'] * len(ids))
        cursor.execute(f"""SELECT item_id, store_location FROM storage
                            WHERE item_id IN ({ph})""", tuple(ids))
        for r in cursor.fetchall():
            current_loc[int(r['item_id'])] = r['store_location'] or ''
    for it in items:
        it['store_location'] = current_loc.get(int(it['item_id']), '')
    units = load_unit_map(cursor)
    conversions = load_item_conversions(cursor, [it['item_id'] for it in items])
    for it in items:
        ordered = float(it.get('qty_ordered') or 0)
        received = float(it.get('qty_received_total') or 0)
        accepted = float(it.get('qty_accepted_total') or 0)
        rejected = float(it.get('qty_rejected_total') or 0)
        # Anything received but not yet passed or failed is still in QC. It may
        # turn out good, so it counts against the order for now.
        awaiting = round(max(0.0, received - accepted - rejected), QTY_DP)
        it['qty_received_total'] = round(received, QTY_DP)
        it['qty_accepted_total'] = round(accepted, QTY_DP)
        it['qty_rejected_total'] = round(rejected, QTY_DP)
        it['qty_awaiting_qc'] = awaiting

        # ORDER OF OPERATIONS MATTERS HERE.
        #
        # qty_ordered is in the unit the ORDER was raised in (kg). Everything
        # from the GRN — received, accepted, rejected — is in the unit the item
        # is STOCKED in (g). Subtracting one from the other directly would take
        # grams away from kilograms, which is the very mistake this whole
        # change exists to prevent.
        #
        # So the ordered quantity is converted into the item's unit FIRST, and
        # only then is anything subtracted from it.
        ordered_conv, status = convert_qty(units, ordered,
                                           it.get('unit_id'), it.get('item_unit_id'),
                                           conversions, it.get('item_id'))
        it['unit_status'] = status
        it['po_unit'] = it.get('unit') or ''
        if status in ('same', 'converted', 'item_factor'):
            ordered_item_unit = round(ordered_conv, QTY_DP)
            it['ordered_in_item_unit'] = ordered_item_unit
            # Rejected quantity does NOT count as fulfilling the order, so a QC
            # rejection reopens the line by that amount without anyone adjusting it.
            it['qty_pending'] = round(max(0.0, ordered_item_unit - accepted - awaiting), QTY_DP)
            it['cap_enforced'] = True
            it['unit_note'] = None if status == 'same' else (
                f"Ordered {_trim_number(ordered)} {it['po_unit']} "
                f"= {_trim_number(ordered_item_unit)} {it.get('item_unit') or ''}")
        else:
            # Nothing sensible to cap against. Per the agreed rule the receipt is
            # still allowed; the screen says why it is unchecked.
            it['qty_pending'] = None
            it['ordered_in_item_unit'] = None
            it['cap_enforced'] = False
            it['unit_note'] = (
                f"Ordered in {it['po_unit'] or '?'}, stocked in "
                f"{it.get('item_unit') or '?'} — not the same kind of measure, "
                f"so the receipt limit cannot be applied"
                if status == 'incompatible' else
                "No stock unit set for this item, so the receipt limit cannot be applied")
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

    # For a GRN raised against a PO, show what was ordered and what is still
    # outstanding. Both are PO-level figures rolled up across EVERY GRN for
    # that PO, not just this one, so a second delivery reads correctly.
    po_ref = None
    po_line_info = {}
    units = load_unit_map(cursor)
    conversions = load_item_conversions(cursor, [g['item_id'] for g in grn_items])
    if grn.get('po_id'):
        cursor.execute("SELECT po_id, po_number FROM purchase_order WHERE po_id = %s",
                       (grn['po_id'],))
        po_ref = cursor.fetchone()

        cursor.execute("""
            SELECT poi.item_id, poi.qty_ordered, poi.unit_id AS po_unit_id,
                   poi.unit_price, i.unit_id AS item_unit_id,
                   COALESCE(r.received, 0) AS received,
                   COALESCE(r.accepted, 0) AS accepted,
                   COALESCE(r.rejected, 0) AS rejected
            FROM po_items poi
            JOIN items i ON i.item_id = poi.item_id
            LEFT JOIN (
                SELECT gi.item_id,
                       SUM(gi.qty_received) AS received,
                       SUM(gi.qty_accepted) AS accepted,
                       SUM(gi.qty_rejected) AS rejected
                FROM grn_item gi
                JOIN grn g ON g.grn_id = gi.grn_id
                WHERE g.po_id = %s
                GROUP BY gi.item_id
            ) r ON r.item_id = poi.item_id
            WHERE poi.po_id = %s
        """, (grn['po_id'], grn['po_id']))
        for r in cursor.fetchall():
            ordered = float(r['qty_ordered'] or 0)
            accepted = float(r['accepted'] or 0)
            received = float(r['received'] or 0)
            awaiting = max(0.0, received - accepted - float(r['rejected'] or 0))
            # Same conversion-before-subtraction rule as the entry form: the
            # ordered figure is in the order's unit, the receipts are in the
            # item's. Outstanding is reported in the ITEM's unit so it lines up
            # with the received quantity shown beside it.
            ordered_conv, o_status = convert_qty(units, ordered,
                                                 r['po_unit_id'], r['item_unit_id'],
                                                 conversions, r['item_id'])
            outstanding = (round(max(0.0, ordered_conv - accepted - awaiting), QTY_DP)
                           if o_status in ('same', 'converted', 'item_factor') else None)
            po_line_info[int(r['item_id'])] = {
                # The order's own figure stays in the order's unit, because
                # that is what the vendor agreed to supply.
                'qty_ordered': round(ordered, QTY_DP),
                'po_received_total': round(received, QTY_DP),
                # Same rule as the GRN entry form: rejected quantity does not
                # count as fulfilling the order, so it reopens the line.
                'po_outstanding': outstanding,
                'outstanding_unit': unit_short(units, r['item_unit_id']),
                'po_unit_id': r['po_unit_id'],
                'item_unit_id': r['item_unit_id'],
                'po_unit': unit_short(units, r['po_unit_id']),
                'unit_price': float(r['unit_price'] or 0),
            }

    for it in grn_items:
        info = po_line_info.get(int(it['item_id']))
        it['po_qty_ordered'] = info['qty_ordered'] if info else None
        it['po_received_total'] = info['po_received_total'] if info else None
        it['po_outstanding'] = info['po_outstanding'] if info else None
        it['outstanding_unit'] = info['outstanding_unit'] if info else None
        it['po_unit'] = info['po_unit'] if info else None
        it['qty_in_po_unit'] = None
        it['line_value'] = None
        if info:
            # The price is per unit of the ORDER, so a receipt booked in the
            # item's unit has to be converted back before it is multiplied.
            # 5000 g against a PO of 10 kg at 500/kg is 5 kg, so 2,500 — not
            # 5000 x 500.
            conv, status = convert_qty(units, float(it.get('qty_received') or 0),
                                       info['item_unit_id'], info['po_unit_id'],
                                       conversions, it.get('item_id'))
            if status in ('same', 'converted', 'item_factor'):
                it['qty_in_po_unit'] = round(conv, QTY_DP)
                it['line_value'] = round(conv * info['unit_price'], 2)

    wo_ref = None
    if grn.get('wo_id'):
        cursor.execute("SELECT wo_number, description FROM work_order WHERE wo_id = %s", (grn['wo_id'],))
        wo_ref = cursor.fetchone()
    cursor.close()
    db.close()
    return render_template('grn_detail.html', grn=grn, grn_items=grn_items,
                           wo_ref=wo_ref, po_ref=po_ref)

# ─── GRN QC INSPECTION ──────────────────────────────────
def _qc_would_exceed_order(cursor, grn_id, grn_item_ids, qty_accepted):
    """Message describing the first line that would push a PO over, else None.

    Checked across EVERY GRN on the purchase order, not just this one, because
    the quantity that pushes it over usually arrived on a different delivery.
    """
    cursor.execute("SELECT po_id FROM grn WHERE grn_id = %s", (grn_id,))
    row = cursor.fetchone()
    po_id = row['po_id'] if row else None
    if not po_id:
        return None          # not against an order; nothing to exceed

    # What is being submitted, per item. A GRN can carry two lines for the
    # same item, so they are summed rather than checked one at a time.
    submitted, line_ids = {}, {}
    for i, gi_id in enumerate(grn_item_ids):
        try:
            new_acc = float(qty_accepted[i] or 0)
        except (TypeError, ValueError):
            continue
        cursor.execute("SELECT item_id FROM grn_item WHERE grn_item_id = %s", (gi_id,))
        gi = cursor.fetchone()
        if not gi:
            continue
        iid = int(gi['item_id'])
        submitted[iid] = submitted.get(iid, 0.0) + new_acc
        line_ids.setdefault(iid, []).append(int(gi_id))
    if not submitted:
        return None

    units = load_unit_map(cursor)
    conversions = load_item_conversions(cursor, list(submitted.keys()))

    for iid, new_total_here in submitted.items():
        ids = line_ids[iid]
        ph = ','.join(['%s'] * len(ids))
        # Accepted on OTHER lines of this PO — every GRN, minus the lines
        # being edited right now, whose stored values are about to be replaced.
        cursor.execute(f"""SELECT COALESCE(SUM(gi.qty_accepted), 0) AS other_accepted
                             FROM grn_item gi
                             JOIN grn g ON g.grn_id = gi.grn_id
                            WHERE g.po_id = %s AND gi.item_id = %s
                              AND gi.grn_item_id NOT IN ({ph})""",
                       tuple([po_id, iid] + ids))
        row = cursor.fetchone()
        other_accepted = float(row['other_accepted'] or 0) if row else 0.0

        cursor.execute("""SELECT poi.qty_ordered, poi.unit_id AS po_unit_id,
                                 i.unit_id AS item_unit_id, i.item_code, i.`desc`
                            FROM po_items poi
                            JOIN items i ON i.item_id = poi.item_id
                           WHERE poi.po_id = %s AND poi.item_id = %s""",
                       (po_id, iid))
        po_line = cursor.fetchone()
        if not po_line:
            continue         # received against the PO but never ordered; not this check's job

        ordered_conv, status = convert_qty(units, float(po_line['qty_ordered'] or 0),
                                           po_line['po_unit_id'], po_line['item_unit_id'],
                                           conversions, iid)
        if status not in ('same', 'converted', 'item_factor'):
            continue         # units do not relate; no honest comparison to make

        would_be = round(other_accepted + new_total_here, QTY_DP)
        if would_be > round(ordered_conv, QTY_DP) + 1e-9:
            unit = unit_short(units, po_line['item_unit_id']) or ''
            headroom = max(0.0, round(ordered_conv - other_accepted, QTY_DP))
            return (f"{po_line['item_code']} — accepting {_trim_number(new_total_here)} {unit} "
                    f"would bring the total accepted on this order to "
                    f"{_trim_number(would_be)} {unit}, against "
                    f"{_trim_number(ordered_conv)} {unit} ordered. "
                    f"{_trim_number(other_accepted)} {unit} has already been accepted on "
                    f"other receipts for this order, so at most "
                    f"{_trim_number(headroom)} {unit} can be accepted here. "
                    f"If the extra quantity is genuinely wanted, raise or amend the "
                    f"purchase order first.")
    return None


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

        # ─── ACCEPTED MAY NEVER EXCEED ORDERED ──────────────────────
        #
        # Outstanding is `ordered - received + rejected`, so rejecting
        # reopens the line and the vendor can send a replacement. That is
        # right. What was missing is the other side of it: once the
        # replacement has arrived, the original rejection must not be
        # reversible, or the same order is fulfilled twice.
        #
        #   PO 1000 g. GRN1 receives 500, accepts 400, rejects 100.
        #   Outstanding reopens to 600. GRN2 receives and accepts 600.
        #   Back on GRN1, accepted is edited 400 -> 500.
        #   Total accepted 1100 against an order for 1000.
        #
        # Receiving is already capped against outstanding. This is the
        # same cap applied at inspection, which is the only other place
        # the accepted total can move.
        over = _qc_would_exceed_order(cursor, grn_id, grn_item_ids, qty_accepted)
        if over:
            cursor.close()
            db.close()
            flash(over, 'error')
            return redirect(url_for('grn_qc', grn_id=grn_id))

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
                        # The quantity moves, and so does the location when the
                        # receipt named one. Without this the goods arrive at the
                        # new rack while the record still says the old one, which
                        # is worse than never having asked.
                        if loc:
                            cursor.execute("""UPDATE storage
                                                 SET physical_availability = physical_availability + %s,
                                                     store_location = %s
                                               WHERE store_id = %s""",
                                           (delta, loc, s_row['store_id']))
                        else:
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

    # The quantity on this screen is in the unit the item is STOCKED in —
    # accepting 5 adds 5 of that unit to storage. Showing it removes the
    # guesswork now that wire is in grams and cable in metres.
    gi_unit = 'gu.unit_short_name' if has_unit_column(cursor, 'grn_item') else 'NULL'
    gi_join = ("LEFT JOIN units gu ON gu.unit_id = gi.unit_id"
               if has_unit_column(cursor, 'grn_item') else "")
    cursor.execute(f"""
        SELECT gi.*, i.item_code, i.`desc` as item_desc,
               u.unit_short_name AS item_unit,
               {gi_unit}         AS received_unit
        FROM grn_item gi
        JOIN items i ON gi.item_id = i.item_id
        LEFT JOIN units u ON u.unit_id = i.unit_id
        {gi_join}
        WHERE gi.grn_id = %s
        ORDER BY gi.grn_item_id
    """, (grn_id,))
    grn_items = cursor.fetchall()

    # If a line was received in one unit and the item now holds another, the
    # figures on screen do not mean what they did. Say so rather than showing
    # today's unit against yesterday's number.
    for gi in grn_items:
        gi['unit_changed'] = bool(gi.get('received_unit')
                                  and gi.get('item_unit')
                                  and gi['received_unit'] != gi['item_unit'])
    
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
        SELECT gi.*, i.item_code, i.`desc` as item_desc,
               u.unit_short_name AS item_unit
        FROM grn_item gi
        JOIN items i ON gi.item_id = i.item_id
        LEFT JOIN units u ON u.unit_id = i.unit_id
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
            'qty_received': round(recv, QTY_DP),
            'qty_accepted': round(acc, QTY_DP),
            'qty_pending': round(pend, QTY_DP),
            'qty_rejected': round(rej, QTY_DP),
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
            SELECT i.item_id, i.item_code, i.`desc` as item_desc, i.mpn, u.unit_short_name as unit,
                   COALESCE(m.mfr_short_name, m.mfr_full_name) as manufacturer_name
            FROM items i
            LEFT JOIN units u ON i.unit_id = u.unit_id
            LEFT JOIN manufacturer m ON i.mfr_id = m.mfr_id
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


# ─── INVENTORY IMPORT ───────────────────────────────────
# Reads a store inventory workbook and brings it into items / manufacturer /
# units / storage. Column matching is deliberately loose: the sheet is a report
# from another system, so extra columns, duplicate headers and odd spellings
# are expected and ignored rather than rejected.

# Columns that must never be taken as the stock figure, whatever else they are
# called — these are movements and money, not a balance.
_QTY_EXCLUDE = ('grn', 'issued', 'issue', 'rate', 'value', 'price', 'amount')

# A hyphen only separates manufacturer from part number when it has whitespace
# beside it. Bare hyphens belong to the part number itself (ACPL-M61L-060E).
_MFR_SPLIT_RE = re.compile(r'^(.*?)(?:\s+-\s*|\s*-\s+)(.*)$', re.S)

# The file's unit words mapped onto the units already in the database.
UNIT_ALIASES = {
    'num': 'pcs', 'nos': 'pcs', 'no': 'pcs', 'nu': 'pcs', 'pcs': 'pcs',
    'pc': 'pcs', 'piece': 'pcs', 'pieces': 'pcs', 'ea': 'pcs', 'each': 'pcs',
    'mts': 'm', 'mtr': 'm', 'mtrs': 'm', 'mt': 'm', 'meter': 'm',
    'meters': 'm', 'metre': 'm', 'metres': 'm', 'm': 'm',
    'kgs': 'kg', 'kg': 'kg', 'kilogram': 'kg', 'kilograms': 'kg',
    'g': 'g', 'gm': 'g', 'gms': 'g', 'gram': 'g', 'grams': 'g',
    'cm': 'cm', 'cms': 'cm',
}


def looks_like_part_number(name):
    """Rough guess at whether a parsed 'manufacturer' is really a part number.

    Part numbers tend to be one token with digits in them (BZX84, M24C04,
    LDI1117); manufacturers tend to be words (Bourns, TE Connectivity) or
    letter-only abbreviations that are genuine companies (AD, IR, FSC). Only a
    hint for the reviewer — never acted on automatically.
    """
    if not name:
        return False
    s = str(name).strip()
    return bool(s) and (' ' not in s) and any(ch.isdigit() for ch in s)


# Things people type when a cell does not apply. These must never become a
# manufacturer or an MPN — 'NA' as a manufacturer would be a permanent piece of
# junk in the master, since nothing downstream would ever question it.
_NOT_APPLICABLE = {
    'na', 'n/a', 'n.a', 'n.a.', 'n a', '-na-', 'na-',
    'nil', 'none', 'null', 'nan', 'not applicable', 'not available',
    'tbd', 'to be decided', 'no mpn', 'xxx',
    '-', '--', '---', '.', '..', '?', '_',
}


# Symbols that turn up in electronics part numbers and descriptions but cannot
# be stored unless every text column is utf8mb4. Folded to their usual ASCII
# spellings, which is how the industry writes them anyway (uF, ohm, +/-).
_ASCII_SYMBOLS = {
    '\u03bc': 'u', '\u00b5': 'u',            # greek mu, micro sign  -> u
    '\u03a9': 'ohm', '\u2126': 'ohm',        # omega                 -> ohm
    '\u00b0': 'deg', '\u00b1': '+/-',
    '\u00d7': 'x', '\u00f7': '/',
    '\u2013': '-', '\u2014': '-', '\u2212': '-',
    '\u2018': "'", '\u2019': "'",
    '\u201c': '"', '\u201d': '"',
    '\u2026': '...', '\u00a0': ' ',
    '\u00bd': '1/2', '\u00bc': '1/4', '\u00be': '3/4',
}


def ascii_fold(value):
    """Makes a cell safe for a non-utf8mb4 column.

    Known symbols are spelled out, accents are stripped to their base letter,
    and anything still outside ASCII is dropped. Returns the text unchanged
    when it is already plain ASCII, which is the overwhelming majority.
    """
    if value is None:
        return None
    s = str(value)
    if s.isascii():
        return s
    for ch, repl in _ASCII_SYMBOLS.items():
        if ch in s:
            s = s.replace(ch, repl)
    if not s.isascii():
        s = unicodedata.normalize('NFKD', s)
        s = s.encode('ascii', 'ignore').decode('ascii')
    return ' '.join(s.split())


# An item code is a part number, not a date and not punctuation. Rows shaped
# like these are spreadsheet furniture — a printed date, a separator line, a
# stray cell below the last part — and importing them creates non-items that
# then sort ahead of every real code in every dropdown.
_DATE_LIKE_CODE = re.compile(r'^\d{1,4}[.\-/]\d{1,2}[.\-/]\d{1,4}$')


def is_usable_item_code(code):
    """False for anything that cannot be a part number."""
    if code is None:
        return False
    s = str(code).strip()
    if not s:
        return False
    if is_not_applicable(s):
        return False
    if _DATE_LIKE_CODE.match(s):
        return False
    # Punctuation or whitespace only — '-', '---', '.', '/'.
    if not any(ch.isalnum() for ch in s):
        return False
    return True


def _norm_space(value):
    """Collapse runs of whitespace so 'A  B' and 'A B' compare equal."""
    return ' '.join(str(value or '').split())


def is_not_applicable(value):
    """True for blanks and the placeholders above. Case and punctuation
    spacing are ignored, so 'N / A' and 'n.a.' both count."""
    if value is None:
        return True
    s = ' '.join(str(value).split())
    if not s:
        return True
    if s.lower() in _NOT_APPLICABLE:
        return True
    # 'N / A' -> 'n/a'
    return s.lower().replace(' ', '') in _NOT_APPLICABLE


def split_manufacturer_mpn(value):
    """'Keltron - SA1011EMAS' -> ('Keltron', 'SA1011EMAS').

    Returns (None, whole) when there is no spaced hyphen, so a part number
    that merely contains hyphens is never mistaken for a manufacturer.

    Either side is dropped if it is a not-applicable placeholder, so 'NA',
    'NA - NA' and 'Murata - NA' give (None, None), (None, None) and
    ('Murata', None) respectively.
    """
    if is_not_applicable(value):
        return None, None
    s = ' '.join(str(value).split())
    m = _MFR_SPLIT_RE.match(s)
    if not m:
        return None, s
    mfr = m.group(1).strip() or None
    mpn = m.group(2).strip() or None
    if is_not_applicable(mfr):
        mfr = None
    if is_not_applicable(mpn):
        mpn = None
    return mfr, mpn


def _clean_header(v):
    return ' '.join(str(v).strip().lower().split()) if v is not None else ''


def parse_inventory_file(filename, file_bytes):
    """Pulls item rows out of a store inventory workbook.

    Returns (rows, meta). Only the first sheet is read. Matching is by keyword
    so the same code copes with 'Item Name' or 'Item Code', and any column it
    does not recognise is simply left alone.
    """
    ext = (filename or '').lower().rsplit('.', 1)[-1] if '.' in (filename or '') else ''
    grid = []

    if ext == 'xls':
        import xlrd
        book = xlrd.open_workbook(file_contents=file_bytes)
        sheet = book.sheet_by_index(0)
        for r in range(sheet.nrows):
            grid.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
        sheet_name = sheet.name
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(values_only=True):
            grid.append(list(row))
        sheet_name = wb.sheetnames[0]

    if not grid:
        return [], {'sheet': sheet_name, 'header_row': None, 'columns': {}}

    # Find the header: the first row in the top 25 that names an item column.
    header_idx, headers = None, []
    for i in range(min(25, len(grid))):
        cells = [_clean_header(v) for v in grid[i]]
        joined = ' | '.join(cells)
        if ('item' in joined or 'code' in joined) and any(
                k in joined for k in ('stock', 'qty', 'quantity', 'stk')):
            header_idx, headers = i, cells
            break
    if header_idx is None:
        header_idx, headers = 0, [_clean_header(v) for v in grid[0]]

    def find(preds, exclude=()):
        """First column whose header matches, skipping excluded words."""
        for idx, h in enumerate(headers):
            if not h or any(x in h for x in exclude):
                continue
            if any(p(h) for p in preds):
                return idx
        return None

    col_code = find([lambda h: 'item' in h and ('name' in h or 'code' in h),
                     lambda h: h in ('item', 'code', 'item code', 'part code')])
    col_desc = find([lambda h: 'desc' in h, lambda h: h == 'item name' and col_code != 0])
    col_unit = find([lambda h: h in ('unit', 'units', 'uom', 'u/m')])
    col_mpn  = find([lambda h: h == 'mpn' or 'part no' in h or 'part number' in h])
    col_loc  = find([lambda h: 'location' in h or h in ('bin', 'rack', 'store location')])
    # Two different things can appear under an HSN heading, and they must not
    # be confused: a CODE ('85331000') is the tax classification itself, an
    # ID ('110') is the originating system's primary key. Both are numeric,
    # so the heading decides which is which rather than the value.
    col_hsn_code = find([lambda h: 'hsn' in h and 'code' in h, lambda h: h == 'hsn'])
    col_hsn_id   = find([lambda h: 'hsn' in h and ('id' in h or 'ref' in h)])
    # Prefer an explicit closing-stock column, then a physical-stock one, then
    # a generic qty — never a GRN/issued/rate/value column.
    col_qty = (find([lambda h: h == 'stock' or 'closing' in h], exclude=_QTY_EXCLUDE)
               or find([lambda h: 'phy' in h and ('stk' in h or 'stock' in h)], exclude=_QTY_EXCLUDE)
               or find([lambda h: 'stock' in h or 'stk' in h], exclude=_QTY_EXCLUDE)
               or find([lambda h: 'qty' in h or 'quantity' in h], exclude=_QTY_EXCLUDE))

    def cell(row, idx):
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        if v is None:
            return None
        s = ascii_fold(str(v)).strip()
        return None if s == '' or s.lower() == 'nan' else s

    rows = []
    skipped_no_code = 0
    skipped_bad_code = []
    for r in range(header_idx + 1, len(grid)):
        raw_row = grid[r]
        code = cell(raw_row, col_code)
        if not code:
            skipped_no_code += 1
            continue
        code = code.lstrip("'").strip()
        if code.endswith('.0') and code[:-2].isdigit():
            code = code[:-2]           # Excel turns long numeric codes into floats
        if not is_usable_item_code(code):
            skipped_bad_code.append(code)
            continue
        qty_raw = cell(raw_row, col_qty)
        try:
            qty = float(qty_raw) if qty_raw is not None else 0.0
        except (TypeError, ValueError):
            qty = 0.0
        mpn_raw = cell(raw_row, col_mpn)
        mfr, mpn = split_manufacturer_mpn(mpn_raw)

        def usable(v):
            """Blanks and 'NA'-style placeholders are treated as empty, so they
            are never written into the item master or the units table."""
            return None if is_not_applicable(v) else v

        rows.append({
            'item_code': code,
            'desc': usable(cell(raw_row, col_desc)),
            'unit_text': usable(cell(raw_row, col_unit)),
            'manufacturer_name': mfr,
            'mpn': mpn,
            # kept so that, if the reviewer says this is not a manufacturer,
            # the whole original value can be restored as the part number
            'mpn_raw': (' '.join(str(mpn_raw).split())
                        if not is_not_applicable(mpn_raw) else None),
            'location': usable(cell(raw_row, col_loc)),
            'qty': round(max(0.0, qty), QTY_DP),
            # Resolved against the hsn table at import time, not here.
            'hsn_code': usable(cell(raw_row, col_hsn_code)),
            'hsn_source_id': usable(cell(raw_row, col_hsn_id)),
        })

    names = {'item code': col_code, 'description': col_desc, 'unit': col_unit,
             'mpn': col_mpn, 'location': col_loc, 'quantity': col_qty,
             'hsn code': col_hsn_code, 'hsn id': col_hsn_id}
    meta = {
        'sheet': sheet_name,
        'header_row': header_idx + 1,
        'columns': {k: (headers[v] if v is not None and v < len(headers) else None)
                    for k, v in names.items()},
        'skipped_no_code': skipped_no_code,
        'skipped_bad_code': skipped_bad_code,
    }
    return rows, meta


# Words that name PACKAGING rather than a measure. A roll of one cable is 100 m
# and another is 500 m, so 'roll' says nothing about how much there is and can
# never convert to anything — which means a PO raised in it has no receipt cap.
# These are not created as units. The item is left without one so that a person
# decides what it is really measured in, rather than the file deciding badly.
_NOT_A_UNIT = {
    'roll', 'rolls', 'set', 'sets', 'reel', 'reels', 'spool', 'spools',
    'pack', 'packs', 'packet', 'packets', 'box', 'boxes', 'carton', 'cartons',
    'bundle', 'bundles', 'bag', 'bags', 'strip', 'strips', 'tray', 'trays',
    'lot', 'lots', 'coil', 'coils', 'sheet', 'sheets',
}


def is_packaging_not_unit(unit_text):
    if unit_text is None:
        return False
    return ' '.join(str(unit_text).split()).lower() in _NOT_A_UNIT


def _resolve_unit(cursor, unit_text, cache, created):
    """unit_id for a unit word, creating the unit if it is genuinely new."""
    if is_not_applicable(unit_text) or is_packaging_not_unit(unit_text):
        return None
    key = unit_text.strip().lower()
    if key in cache:
        return cache[key]
    short = UNIT_ALIASES.get(key, unit_text.strip())
    cursor.execute("SELECT unit_id FROM units WHERE LOWER(unit_short_name) = %s LIMIT 1",
                   (short.lower(),))
    row = cursor.fetchone()
    if row:
        cache[key] = row['unit_id']
        return cache[key]
    cache[key] = _insert_with_defaults(
        cursor, 'units',
        {'unit_name': short, 'unit_short_name': short},
        # 1/1 means "no conversion". base_unit_id is left NULL on purpose:
        # the file does not say which family a new unit belongs to, and
        # guessing wrong is worse than leaving it to be set by hand.
        overrides={'conv_fact_num': 1, 'conv_fact_denom': 1,
                   'conv_fact_den': 1, 'decimal_places': 3},
    )
    created.append(short)
    return cache[key]


def _resolve_manufacturer(cursor, name, cache, created):
    # Guarded here as well as at the split, so no future caller can create an
    # 'NA' or '-' manufacturer by accident.
    if is_not_applicable(name):
        return None
    key = name.strip().lower()
    if key in cache:
        return cache[key]
    cursor.execute("""SELECT mfr_id FROM manufacturer
                      WHERE LOWER(mfr_short_name) = %s OR LOWER(mfr_full_name) = %s
                      LIMIT 1""", (key, key))
    row = cursor.fetchone()
    if row:
        cache[key] = row['mfr_id']
        return cache[key]
    cache[key] = _insert_with_defaults(
        cursor, 'manufacturer',
        {'mfr_short_name': name.strip(), 'mfr_full_name': name.strip()},
    )
    created.append(name.strip())
    return cache[key]


def hsn_lookup(cursor):
    """Both ways a spreadsheet can name an HSN row, resolved to its hsn_id.

      by_code   '85331000' -> 14    the tax classification itself
      by_source  110       -> 14    the id it carried in the system it came from

    The second only works for rows imported with a source id, which is why the
    HSN import keeps it. Built once per import rather than queried per row.
    """
    have = {c['COLUMN_NAME'] for c in get_table_columns(cursor, 'hsn')}
    fields = 'hsn_id, hsn_code' + (', source_id' if 'source_id' in have else '')
    cursor.execute("SELECT %s FROM hsn WHERE COALESCE(is_deleted, 0) = 0 ORDER BY hsn_id"
                   % fields)
    by_code, by_source = {}, {}
    for row in cursor.fetchall():
        code = str(row['hsn_code'] or '').strip()
        if code and code not in by_code:
            by_code[code] = int(row['hsn_id'])
        src = row.get('source_id')
        if src is not None and int(src) not in by_source:
            by_source[int(src)] = int(row['hsn_id'])
    return by_code, by_source


def resolve_row_hsn(row, by_code, by_source):
    """(hsn_id, what the file said). Both None when the file said nothing.

    A code is preferred over an id when the file carries both, because the code
    means the same thing in every system and the id does not.
    """
    code = str(row.get('hsn_code') or '').strip()
    if code:
        return by_code.get(code), code

    ref = str(row.get('hsn_source_id') or '').strip()
    if not ref:
        return None, None
    try:
        n = int(float(ref))
    except (TypeError, ValueError):
        return None, ref
    if n == 0:
        return None, None      # 'no HSN', written as a zero rather than left blank
    return by_source.get(n), ref


def preview_inventory_import(cursor, rows):
    """Works out what the import would do, without changing anything.

    Given the size of these files (thousands of rows, dozens of new
    manufacturers) the effects are worth seeing before they happen.
    """
    seen_codes = {}
    for r in rows:
        seen_codes.setdefault(r['item_code'], 0)
        seen_codes[r['item_code']] += 1
    duplicates = {k: v for k, v in seen_codes.items() if v > 1}

    codes = list(seen_codes.keys())
    existing_items = {}
    for i in range(0, len(codes), 500):
        chunk = codes[i:i + 500]
        ph = ','.join(['%s'] * len(chunk))
        cursor.execute(f"""SELECT item_id, item_code, unit_id, mfr_id, `desc`, mpn
                           FROM items WHERE item_code IN ({ph})""", tuple(chunk))
        for row in cursor.fetchall():
            existing_items[row['item_code']] = row

    # An item existing in the master is NOT the same as that item holding
    # stock. After the store is cleared every item still exists, so counting
    # 'existing items' would wrongly warn that thousands of quantities are
    # about to be doubled. Ask storage directly.
    items_with_stock = 0
    qty_already_held = 0.0
    existing_ids = [row['item_id'] for row in existing_items.values()]
    for i in range(0, len(existing_ids), 500):
        chunk = existing_ids[i:i + 500]
        ph = ','.join(['%s'] * len(chunk))
        cursor.execute(f"""SELECT COUNT(*) AS n, COALESCE(SUM(physical_availability), 0) AS q
                           FROM storage
                           WHERE item_id IN ({ph}) AND physical_availability > 0""",
                       tuple(chunk))
        row = cursor.fetchone()
        if row:
            items_with_stock += int(row['n'] or 0)
            qty_already_held += float(row['q'] or 0)

    cursor.execute("""SELECT mfr_id, mfr_short_name, mfr_full_name,
                             LOWER(mfr_short_name) AS s, LOWER(mfr_full_name) AS f
                        FROM manufacturer""")
    known_mfrs = set()
    mfr_name_by_id = {}
    mfr_id_by_name = {}
    for row in cursor.fetchall():
        known_mfrs.add(row['s'] or '')
        known_mfrs.add(row['f'] or '')
        mfr_name_by_id[row['mfr_id']] = row['mfr_short_name'] or row['mfr_full_name'] or ''
        for key in (row['s'], row['f']):
            if key:
                mfr_id_by_name.setdefault(key, row['mfr_id'])
    cursor.execute("SELECT LOWER(unit_short_name) AS s FROM units")
    known_units = {row['s'] for row in cursor.fetchall()}

    # Keyed by lowercase name so 'Vishay' and 'vishay' count once, matching
    # what the import itself does — a preview that overstates is misleading.
    # Each entry carries how many rows use it and one raw cell as evidence,
    # because the reviewer cannot judge 'BZX84' without seeing 'BZX84 - C5V1'.
    new_mfrs, new_units, packaging = {}, set(), {}
    # What the file would OVERWRITE on items that already exist. Worth seeing
    # before committing: these replace a value somebody may have corrected by
    # hand, and there is no undo short of the backup.
    desc_overwrites, mfr_overwrites = [], []
    for r in rows:
        existing = existing_items.get(r['item_code'])
        if existing:
            file_desc = (r['desc'] or '').strip()[:100]
            old_desc = (existing.get('desc') or '').strip()
            if file_desc and old_desc and _norm_space(file_desc) != _norm_space(old_desc):
                desc_overwrites.append({'item_code': r['item_code'],
                                        'was': old_desc, 'now': file_desc})
            if r['manufacturer_name'] and existing.get('mfr_id'):
                file_mfr_id = mfr_id_by_name.get(r['manufacturer_name'].strip().lower())
                if file_mfr_id and file_mfr_id != existing['mfr_id']:
                    mfr_overwrites.append({
                        'item_code': r['item_code'],
                        'was': mfr_name_by_id.get(existing['mfr_id'], '?'),
                        'now': r['manufacturer_name'].strip()})
        if r['manufacturer_name']:
            key = r['manufacturer_name'].strip().lower()
            if key not in known_mfrs:
                entry = new_mfrs.setdefault(key, {
                    'name': r['manufacturer_name'].strip(),
                    'count': 0,
                    'sample': r.get('mpn_raw') or '',
                    'suspicious': looks_like_part_number(r['manufacturer_name']),
                })
                entry['count'] += 1
        if r['unit_text']:
            if is_packaging_not_unit(r['unit_text']):
                # Counted separately: these are skipped rather than created,
                # and the reviewer should know which rows lose their unit.
                word = ' '.join(str(r['unit_text']).split())
                packaging[word] = packaging.get(word, 0) + 1
                continue
            short = UNIT_ALIASES.get(r['unit_text'].strip().lower(), r['unit_text'].strip())
            if short.lower() not in known_units:
                new_units.add(short)

    # What the HSN column in this file would do. Shown before the import so a
    # file whose ids mean nothing in this database is caught here rather
    # than discovered later as items with no tax rate.
    hsn_by_code, hsn_by_source = hsn_lookup(cursor)
    hsn_ok, hsn_bad, hsn_named = 0, {}, 0
    for r in rows:
        hid, stated = resolve_row_hsn(r, hsn_by_code, hsn_by_source)
        if not stated:
            continue
        hsn_named += 1
        if hid:
            hsn_ok += 1
        else:
            hsn_bad[stated] = hsn_bad.get(stated, 0) + 1

    return {
        'total_rows': len(rows),
        'distinct_items': len(seen_codes),
        'hsn_rows_named': hsn_named,
        'hsn_rows_resolved': hsn_ok,
        'hsn_unresolved': dict(sorted(hsn_bad.items(),
                                      key=lambda kv: -kv[1])[:40]),
        'hsn_unresolved_total': sum(hsn_bad.values()),
        'existing_items': len(existing_items),
        'new_items': len(seen_codes) - len(existing_items),
        'items_with_stock': items_with_stock,
        'qty_already_held': round(qty_already_held, QTY_DP),
        'duplicates': duplicates,
        # suspicious first, so the ones needing a decision are at the top
        'new_manufacturers': sorted(new_mfrs.values(),
                                    key=lambda e: (not e['suspicious'], e['name'].lower())),
        'new_units': sorted(new_units),
        # {word: row count} for packaging words that will NOT become units
        'packaging_units': dict(sorted(packaging.items())),
        'total_qty': round(sum(r['qty'] for r in rows), QTY_DP),
        'rows_with_qty': sum(1 for r in rows if r['qty'] > 0),
        'desc_overwrites': desc_overwrites,
        'mfr_overwrites': mfr_overwrites,
    }


def apply_inventory_import(cursor, rows, default_location='Main Store',
                           not_manufacturers=None):
    """Writes the parsed rows into items / manufacturer / units / storage.

    Stock is ADDED to whatever the item already holds, so the file is treated
    as a receipt rather than a stocktake. Importing the same file twice will
    therefore add the quantities twice.

    not_manufacturers holds the lower-cased names the reviewer rejected on the
    preview screen. For those rows the hyphen was part of the part number, not
    a manufacturer prefix, so the original MPN cell is restored untouched.
    """
    rejected = {str(n).strip().lower() for n in (not_manufacturers or ()) if str(n).strip()}
    mfr_cache, unit_cache = {}, {}
    hsn_by_code, hsn_by_source = hsn_lookup(cursor)
    hsn_linked, hsn_rechanged, hsn_unresolved = 0, [], {}
    created_mfrs, created_units, created_items = [], [], []
    desc_changed, mfr_changed = [], []
    items_updated = 0
    storage_rows_created = 0

    for r in rows:
        mfr_name, mpn = r['manufacturer_name'], r['mpn']
        if mfr_name and mfr_name.strip().lower() in rejected:
            # Not a manufacturer after all — put the cell back the way it came.
            mfr_name = None
            mpn = r.get('mpn_raw') or mpn

        mfr_id = _resolve_manufacturer(cursor, mfr_name, mfr_cache, created_mfrs)
        unit_id = _resolve_unit(cursor, r['unit_text'], unit_cache, created_units)

        cursor.execute("SELECT * FROM items WHERE item_code = %s LIMIT 1", (r['item_code'],))
        item = cursor.fetchone()

        if item:
            item_id = item['item_id']
            sets, params = [], []

            # DESCRIPTION and MANUFACTURER follow the file. The store edits
            # these in the spreadsheet and expects the correction to arrive,
            # so a newer value replaces the old one.
            #
            # A BLANK cell never blanks the master. Only a value the file
            # actually states can change anything — otherwise a column left
            # empty on one export would wipe descriptions wholesale.
            new_desc = (r['desc'] or '').strip()[:100]
            old_desc = (item.get('desc') or '').strip()
            if new_desc and _norm_space(new_desc) != _norm_space(old_desc):
                sets.append("`desc` = %s"); params.append(new_desc)
                if old_desc:
                    desc_changed.append({'item_code': r['item_code'],
                                         'was': old_desc, 'now': new_desc})

            if mfr_id and mfr_id != item.get('mfr_id'):
                sets.append("mfr_id = %s"); params.append(mfr_id)
                if item.get('mfr_id'):
                    mfr_changed.append({'item_code': r['item_code'],
                                        'now': (mfr_name or '').strip()})

            # MPN and UNIT stay gap-fill. Unit especially: the twenty-one
            # copper wire items were moved to grams by hand and the file still
            # calls them kilograms, metres and pieces. Letting the file win
            # would undo that on every import, and silently.
            if mpn and not (item.get('mpn') or '').strip():
                sets.append("mpn = %s"); params.append(mpn[:100])
            if unit_id and not item.get('unit_id'):
                sets.append("unit_id = %s"); params.append(unit_id)

            if item.get('is_deleted'):
                sets.append("is_deleted = 0")
            if sets:
                params.append(item_id)
                cursor.execute(f"UPDATE items SET {', '.join(sets)} WHERE item_id = %s",
                               tuple(params))
                items_updated += 1
        else:
            columns = get_table_columns(cursor, 'items')
            col_names = {c['COLUMN_NAME'] for c in columns}
            values = {'item_code': r['item_code']}
            if 'desc' in col_names:
                values['desc'] = (r['desc'] or r['item_code'])[:100]
            if 'mpn' in col_names and mpn:
                values['mpn'] = mpn[:100]
            if 'mfr_id' in col_names and mfr_id:
                values['mfr_id'] = mfr_id
            if 'unit_id' in col_names and unit_id:
                values['unit_id'] = unit_id
            if 'is_deleted' in col_names:
                values['is_deleted'] = 0
            item_id = _insert_with_defaults(cursor, 'items', values)
            created_items.append(r['item_code'])

        # The HSN lives on the item, in hsn_items, one per item. The file
        # names it either by code or by the id it had in the system it came
        # from; both end at the same hsn_id. A blank cell changes nothing,
        # and a value that cannot be resolved is counted and reported rather
        # than written as something else.
        hsn_id, stated = resolve_row_hsn(r, hsn_by_code, hsn_by_source)
        if stated and hsn_id:
            cursor.execute("""SELECT hsn_id FROM hsn_items WHERE item_id = %s""",
                           (item_id,))
            was = cursor.fetchone()
            if not was:
                cursor.execute("""INSERT INTO hsn_items (item_id, hsn_id)
                                  VALUES (%s, %s)
                                  ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)""",
                               (item_id, hsn_id))
                hsn_linked += 1
            elif int(was['hsn_id']) != hsn_id:
                cursor.execute("UPDATE hsn_items SET hsn_id = %s WHERE item_id = %s",
                               (hsn_id, item_id))
                hsn_rechanged.append({'item_code': r['item_code'], 'now': stated})
        elif stated:
            hsn_unresolved[stated] = hsn_unresolved.get(stated, 0) + 1

        # storage holds one row per item (uq_storage_item), so add to it
        location = r['location'] or default_location
        cursor.execute("SELECT store_id, physical_availability FROM storage WHERE item_id = %s LIMIT 1",
                       (item_id,))
        s_row = cursor.fetchone()
        if s_row:
            cursor.execute("""UPDATE storage
                              SET physical_availability = physical_availability + %s,
                                  store_location = COALESCE(NULLIF(%s, ''), store_location)
                              WHERE store_id = %s""",
                           (r['qty'], location, s_row['store_id']))
        else:
            cursor.execute("""INSERT INTO storage (item_id, store_location, physical_availability, created_at)
                              VALUES (%s, %s, %s, NOW())""",
                           (item_id, location, r['qty']))
            storage_rows_created += 1

    return {
        'rows': len(rows),
        'items_created': len(created_items),
        'items_updated': items_updated,
        'manufacturers_created': created_mfrs,
        'units_created': created_units,
        'storage_rows_created': storage_rows_created,
        'manufacturers_rejected': sorted(rejected),
        'desc_changed': desc_changed,
        'mfr_changed': mfr_changed,
        'hsn_linked': hsn_linked,
        'hsn_changed': hsn_rechanged,
        'hsn_unresolved': dict(sorted(hsn_unresolved.items())),
    }


# ═══════════════════════════════════════════════════════════
#  VENDOR IMPORT
# ═══════════════════════════════════════════════════════════

# Spreadsheet heading -> vendor column. Matching is by keyword so the same
# code copes with 'GST No', 'gst_no' and 'GSTIN'. A heading that matches
# nothing is reported rather than silently dropped.
_VENDOR_COLUMNS = [
    ('short_name',     lambda h: h in ('short name', 'short_name', 'shortname',
                                       'vendor', 'vendor name', 'name', 'code')),
    ('full_name',      lambda h: 'full' in h and 'name' in h or h in ('legal name', 'company')),
    ('gst_no',         lambda h: 'gst' in h),
    # Four free-text lines rather than named parts. city, pincode and
    # country were dropped from the table, so a file still carrying those
    # headings will have them listed as unrecognised rather than silently
    # thrown away — which is the honest outcome, and visible on the preview.
    ('address_line_1', lambda h: h in ('address_line_1', 'address line 1', 'address1',
                                       'address', 'addr')),
    ('address_line_2', lambda h: h in ('address_line_2', 'address line 2', 'address2')),
    ('address_line_3', lambda h: h in ('address_line_3', 'address line 3', 'address3')),
    ('address_line_4', lambda h: h in ('address_line_4', 'address line 4', 'address4')),
    # The table has been through two address shapes. Both sets of headings
    # are recognised; the import writes only the columns that exist, so a
    # file carrying either is handled without editing it.
    ('pincode',        lambda h: 'pin' in h or h in ('zip', 'postcode', 'postal code')),
    ('city',           lambda h: h in ('city', 'town', 'district')),
    ('state',          lambda h: h == 'state'),
    ('country',        lambda h: h == 'country'),
    ('ph_no',          lambda h: h in ('ph_no', 'ph no', 'phone', 'phone no',
                                       'mobile', 'contact', 'contact no', 'telephone')),
    ('email_id',       lambda h: 'email' in h or h in ('mail', 'e-mail', 'email id')),
    # How this vendor's tax is split. The HSN code says the RATE; this says
    # whether that rate is charged as CGST+SGST, as IGST, or not at all.
    #   1  within the state      -> CGST + SGST, half each
    #   2  another state         -> IGST, the whole rate
    #   0  outside India         -> neither
    ('tax_mode',       lambda h: h in ('tax_mode', 'tax mode', 'taxmode', 'gst_mode',
                                       'gst mode', 'tax type')),
    # Which currency this vendor invoices in. The spreadsheet carries the
    # id, not the code, and those ids line up with the currency table:
    #   1 INR   2 USD   3 EUR   4 AED   5 MYR
    # The preview resolves each one to its code so a mismatch is seen
    # before it is imported rather than after.
    ('currency_id',    lambda h: h in ('currency_id', 'currency id', 'currencyid',
                                       'currency', 'curr_id', 'curr id', 'curr code')),
]


def parse_vendor_file(filename, file_bytes):
    """Pull vendor rows out of a spreadsheet.

    Returns (rows, meta). The header is found by looking for a row that names
    a vendor column, so a title line or a blank row above it does no harm —
    the sample file has a blank row between the header and the first vendor.
    """
    ext = (filename or '').lower().rsplit('.', 1)[-1] if '.' in (filename or '') else ''
    grid = []
    if ext == 'xls':
        import xlrd
        book = xlrd.open_workbook(file_contents=file_bytes)
        sheet = book.sheet_by_index(0)
        for r in range(sheet.nrows):
            grid.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
        sheet_name = sheet.name
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(values_only=True):
            grid.append(list(row))
        sheet_name = wb.sheetnames[0]

    if not grid:
        return [], {'sheet': sheet_name, 'header_row': None, 'columns': {},
                    'unmapped_headings': [], 'skipped_no_name': 0}

    header_idx, headers = None, []
    for i in range(min(25, len(grid))):
        cells = [_clean_header(v) for v in grid[i]]
        joined = ' | '.join(cells)
        if 'name' in joined or 'vendor' in joined or 'gst' in joined:
            header_idx, headers = i, cells
            break
    if header_idx is None:
        header_idx, headers = 0, [_clean_header(v) for v in grid[0]]

    col_of, used = {}, set()
    for field, pred in _VENDOR_COLUMNS:
        for idx, h in enumerate(headers):
            if not h or idx in used:
                continue
            if pred(h):
                col_of[field] = idx
                used.add(idx)
                break
    unmapped = [headers[i] for i in range(len(headers))
                if headers[i] and i not in used]

    def cell(row, field):
        idx = col_of.get(field)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        if v is None:
            return None
        # openpyxl hands back numbers for pincode and phone; a phone number is
        # a label, not a quantity, and 8696869000.0 is not a phone number.
        if isinstance(v, float) and v == int(v):
            v = int(v)
        s = ascii_fold(str(v)).strip()
        return None if s == '' or s.lower() == 'nan' else s

    rows, skipped = [], 0
    for r in range(header_idx + 1, len(grid)):
        raw = grid[r]
        rec = {field: cell(raw, field) for field, _ in _VENDOR_COLUMNS}
        # A vendor with no name is not a vendor. Blank rows and totals go here.
        if is_not_applicable(rec.get('short_name')) and is_not_applicable(rec.get('full_name')):
            skipped += 1
            continue
        for k, v in list(rec.items()):
            if is_not_applicable(v):
                rec[k] = None
        # tax_mode is a code, not text. 0 is meaningful here — it is the
        # international mode — so it must not be treated as an empty cell.
        if rec.get('tax_mode') is not None:
            try:
                rec['tax_mode'] = int(float(str(rec['tax_mode']).strip()))
            except (TypeError, ValueError):
                rec['tax_mode'] = None
        # Same for the currency id. 0 is not a currency, so it is dropped
        # rather than written as one.
        if rec.get('currency_id') is not None:
            try:
                rec['currency_id'] = int(float(str(rec['currency_id']).strip())) or None
            except (TypeError, ValueError):
                rec['currency_id'] = None
        # The file leaves full_name blank for every row in the sample. The
        # short name is the only name there is, so it stands for both rather
        # than leaving a NOT NULL column empty.
        if not rec.get('short_name'):
            rec['short_name'] = rec['full_name']
        if not rec.get('full_name'):
            rec['full_name'] = rec['short_name']
        rows.append(rec)

    return rows, {
        'sheet': sheet_name,
        'header_row': header_idx + 1,
        'columns': {f: (headers[i] if i < len(headers) else None)
                    for f, i in col_of.items()},
        'unmapped_headings': unmapped,
        'skipped_no_name': skipped,
    }


def preview_vendor_import(cursor, rows):
    """What the file would do, without doing it."""
    cols = {c['COLUMN_NAME'] for c in get_table_columns(cursor, 'vendor')}

    names = [r['short_name'] for r in rows if r.get('short_name')]
    existing = {}
    if names:
        ph = ','.join(['%s'] * len(names))
        cursor.execute(f"""SELECT vendor_id, short_name, full_name, gst_no, is_deleted
                             FROM vendor WHERE short_name IN ({ph})""", tuple(names))
        for row in cursor.fetchall():
            existing[(row['short_name'] or '').strip().lower()] = row

    # A GST number belongs to one company. The same number under two names
    # means one of them is wrong, and it is worth saying so before importing.
    gsts = [r['gst_no'] for r in rows if r.get('gst_no')]
    gst_clash = []
    if gsts and 'gst_no' in cols:
        ph = ','.join(['%s'] * len(gsts))
        cursor.execute(f"""SELECT vendor_id, short_name, gst_no
                             FROM vendor WHERE gst_no IN ({ph}) AND is_deleted = 0""",
                       tuple(gsts))
        by_gst = {(r['gst_no'] or '').strip().upper(): r for r in cursor.fetchall()}
        for r in rows:
            g = (r.get('gst_no') or '').strip().upper()
            hit = by_gst.get(g)
            if hit and (hit['short_name'] or '').strip().lower() != (r['short_name'] or '').strip().lower():
                gst_clash.append({'gst_no': r['gst_no'],
                                  'in_file': r['short_name'],
                                  'on_record': hit['short_name']})

    seen, dupes = {}, {}
    for r in rows:
        k = (r.get('short_name') or '').strip().lower()
        seen[k] = seen.get(k, 0) + 1
    dupes = {k: v for k, v in seen.items() if v > 1 and k}

    updates, creates = [], []
    for r in rows:
        k = (r.get('short_name') or '').strip().lower()
        if k in existing:
            changes = []
            for field in ('full_name', 'gst_no'):
                if field in cols and r.get(field):
                    was = (existing[k].get(field) or '').strip()
                    if _norm_space(r[field]) != _norm_space(was):
                        changes.append({'field': field, 'was': was or '—', 'now': r[field]})
            updates.append({'short_name': r['short_name'], 'changes': changes,
                            'revived': bool(existing[k].get('is_deleted'))})
        else:
            creates.append(r['short_name'])

    # An id that is not in the currency table would be refused by the foreign
    # key, and an id that resolves to the wrong code is worse — it imports
    # cleanly and prices in the wrong money. Both are shown here.
    currency_use, bad_currency = [], []
    # Only worth summarising when there is a column to write it to.
    # Otherwise the no-column warning below is the honest message, and a
    # tidy breakdown of currencies would imply they are being stored.
    if 'currency_id' in cols and any(r.get('currency_id') for r in rows):
        cursor.execute("""SELECT curr_id, curr_code, curr_name FROM currency
                           WHERE COALESCE(curr_is_deleted, 0) = 0""")
        known = {int(c['curr_id']): c for c in cursor.fetchall()}
        counts = {}
        for r in rows:
            cid = r.get('currency_id')
            if not cid:
                continue
            counts[int(cid)] = counts.get(int(cid), 0) + 1
        for cid, n in sorted(counts.items()):
            c = known.get(cid)
            if c:
                currency_use.append({'currency_id': cid, 'code': c['curr_code'],
                                     'name': c['curr_name'], 'vendors': n})
            else:
                bad_currency.append({'currency_id': cid, 'vendors': n})

    # Columns the file has that the table does not. Reported, not dropped in
    # silence: somebody typed them for a reason.
    file_fields = {f for f, _ in _VENDOR_COLUMNS if any(r.get(f) for r in rows)}
    no_column = sorted(file_fields - cols)

    return {
        'total_rows': len(rows),
        'new_vendors': creates,
        'existing_vendors': updates,
        'duplicates_in_file': dupes,
        'gst_clashes': gst_clash,
        'fields_with_no_column': no_column,
        'currency_use': currency_use,
        'unknown_currency_ids': bad_currency,
        'table_columns': sorted(cols),
    }


def apply_vendor_import(cursor, rows):
    """Write the rows. Matched on short_name; blanks never blank a value."""
    columns = get_table_columns(cursor, 'vendor')
    cols = {c['COLUMN_NAME'] for c in columns}
    # _insert_with_defaults trims to the column width on INSERT. The UPDATE
    # path below has to do the same, or a long address would import fine as a
    # new vendor and abort as an existing one — the same file behaving two
    # different ways depending on what is already in the table.
    widths = {c['COLUMN_NAME']: c.get('CHARACTER_MAXIMUM_LENGTH') for c in columns}
    created, updated, revived, truncated = [], [], [], []

    for r in rows:
        name = (r.get('short_name') or '').strip()
        if not name:
            continue
        cursor.execute("SELECT * FROM vendor WHERE short_name = %s LIMIT 1", (name,))
        existing = cursor.fetchone()

        payload = {f: v for f, v in r.items() if f in cols and v is not None}
        for f, v in list(payload.items()):
            w = widths.get(f)
            if w and isinstance(v, str) and len(v) > w:
                payload[f] = v[:w]
                truncated.append({'short_name': name, 'field': f,
                                  'kept': w, 'was': len(v)})

        if existing:
            sets, params = [], []
            for field, value in payload.items():
                if field == 'short_name':
                    continue                      # the key it was matched on
                was = existing.get(field)
                if _norm_space(value) != _norm_space(was or ''):
                    sets.append(f"`{field}` = %s")
                    params.append(value)
            if existing.get('is_deleted'):
                sets.append("is_deleted = 0")
                revived.append(name)
            if sets:
                params.append(existing['vendor_id'])
                cursor.execute(f"UPDATE vendor SET {', '.join(sets)} WHERE vendor_id = %s",
                               tuple(params))
                updated.append(name)
        else:
            if 'is_deleted' in cols:
                payload['is_deleted'] = 0
            _insert_with_defaults(cursor, 'vendor', payload)
            created.append(name)

    return {'rows': len(rows), 'created': created,
            'updated': updated, 'revived': revived, 'truncated': truncated}


# ═══════════════════════════════════════════════════════════
#  HSN IMPORT
# ═══════════════════════════════════════════════════════════

_HSN_COLUMNS = [
    ('hsn_code',      lambda h: 'hsn' in h and ('code' in h or h == 'hsn') or h == 'code'),
    ('description',   lambda h: 'desc' in h or h in ('particulars', 'goods', 'details')),
    # The id this row had in the system it came from. Worth keeping: the item
    # file refers to HSN rows by that id, not by the code, so without it the
    # two files cannot be joined once they are in the database.
    ('source_id',     lambda h: h in ('id', 'source_id', 'source id', 'hsn_id',
                                      'hsnid', 'hsn ref', 'ref id', 'trippsy id')),
    ('tax_rate',      lambda h: h in ('tax_rate','tax rate','rate','gst','gst rate',
                                      'gst%','total','total rate','tax')),
    ('cgst',          lambda h: h.startswith('cgst')),
    ('sgst',          lambda h: h.startswith('sgst')),
    ('igst',          lambda h: h.startswith('igst')),
    ('tax_date_from', lambda h: ('from' in h and 'date' in h) or h in ('effective from',
                                 'w.e.f', 'wef', 'tax_date_from', 'valid from')),
    ('tax_date_to',   lambda h: ('to' in h and 'date' in h) or h in ('tax_date_to',
                                 'valid to', 'effective to')),
    # Optional. When present the HSN is also linked to that item, which is
    # how the rest of the application finds an item's tax rate.
    ('item_code',     lambda h: h in ('item_code','item code','item','item name','part code')),
]


def _as_rate(value):
    """A tax rate from a cell. '18%', '18.00', 18 and '18 %' all give 18.0."""
    if value is None:
        return None
    s = str(value).strip().replace('%', '').replace(',', '').strip()
    if not s:
        return None
    try:
        return round(float(s), 2)
    except (TypeError, ValueError):
        return None


def _as_int(value):
    """A whole number from a cell, or None. Excel hands back 110.0 for 110."""
    if value is None:
        return None
    s = str(value).strip()
    if not s or is_not_applicable(s):
        return None
    try:
        return int(float(s))
    except (TypeError, ValueError):
        return None


def _as_date(value):
    """A date from a cell, or None. Excel hands back datetimes; people type text."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s or is_not_applicable(s):
        return None
    s = s.split(' ')[0]
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y',
                '%d.%m.%Y', '%Y/%m/%d', '%d-%b-%Y', '%d %b %Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_hsn_file(filename, file_bytes):
    """Pull HSN rows out of a spreadsheet. Returns (rows, meta)."""
    ext = (filename or '').lower().rsplit('.', 1)[-1] if '.' in (filename or '') else ''
    grid = []
    if ext == 'xls':
        import xlrd
        book = xlrd.open_workbook(file_contents=file_bytes)
        sheet = book.sheet_by_index(0)
        for r in range(sheet.nrows):
            grid.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
        sheet_name = sheet.name
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(values_only=True):
            grid.append(list(row))
        sheet_name = wb.sheetnames[0]

    if not grid:
        return [], {'sheet': sheet_name, 'header_row': None, 'columns': {},
                    'unmapped_headings': [], 'skipped_no_code': 0}

    header_idx, headers = None, []
    for i in range(min(25, len(grid))):
        cells = [_clean_header(v) for v in grid[i]]
        if 'hsn' in ' | '.join(cells):
            header_idx, headers = i, cells
            break
    if header_idx is None:
        header_idx, headers = 0, [_clean_header(v) for v in grid[0]]

    col_of, used = {}, set()
    for field, pred in _HSN_COLUMNS:
        for idx, h in enumerate(headers):
            if not h or idx in used:
                continue
            if pred(h):
                col_of[field] = idx
                used.add(idx)
                break
    unmapped = [headers[i] for i in range(len(headers)) if headers[i] and i not in used]

    def raw(row, field):
        idx = col_of.get(field)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    def text(row, field):
        v = raw(row, field)
        if v is None:
            return None
        if isinstance(v, float) and v == int(v):
            v = int(v)            # 853400.0 is not an HSN code
        s = ascii_fold(str(v)).strip()
        return None if s == '' or s.lower() == 'nan' or is_not_applicable(s) else s

    rows, skipped = [], 0
    for r in range(header_idx + 1, len(grid)):
        line = grid[r]
        code = text(line, 'hsn_code')
        if not code:
            skipped += 1
            continue
        code = code.lstrip("'").strip()

        rate = _as_rate(raw(line, 'tax_rate'))
        cgst = _as_rate(raw(line, 'cgst'))
        sgst = _as_rate(raw(line, 'sgst'))
        igst = _as_rate(raw(line, 'igst'))

        # The eight HSN codes already seeded in this database all carry
        # tax_rate 18, cgst 9, sgst 9, igst 18 — the split for a sale inside
        # the state and the single rate for one across it, both held against
        # the same code. A file that states only the total is completed that
        # way rather than left half empty.
        if rate is not None:
            if cgst is None and sgst is None:
                cgst = sgst = round(rate / 2.0, 2)
            if igst is None:
                igst = rate
        elif cgst is not None or sgst is not None:
            rate = round((cgst or 0) + (sgst or 0), 2)
            if igst is None:
                igst = rate
        elif igst is not None:
            rate = igst
            cgst = sgst = round(igst / 2.0, 2)

        rows.append({
            'hsn_code':      code[:50],
            'description':   (text(line, 'description') or None),
            'tax_rate':      rate,
            'cgst':          cgst,
            'sgst':          sgst,
            'igst':          igst,
            'tax_date_from': _as_date(raw(line, 'tax_date_from')),
            'tax_date_to':   _as_date(raw(line, 'tax_date_to')),
            'item_code':     text(line, 'item_code'),
            'source_id':     _as_int(raw(line, 'source_id')),
        })

    return rows, {
        'sheet': sheet_name,
        'header_row': header_idx + 1,
        'columns': {f: (headers[i] if i < len(headers) else None) for f, i in col_of.items()},
        'unmapped_headings': unmapped,
        'skipped_no_code': skipped,
    }


def preview_hsn_import(cursor, rows):
    codes = [r['hsn_code'] for r in rows if r['hsn_code']]
    existing, ambiguous = {}, []
    if codes:
        ph = ','.join(['%s'] * len(codes))
        cursor.execute(f"""SELECT hsn_id, hsn_code, description, tax_rate, cgst, sgst,
                                  igst, tax_date_from, is_deleted
                             FROM hsn
                            WHERE hsn_code IN ({ph}) AND COALESCE(is_deleted,0) = 0
                            ORDER BY hsn_id""", tuple(codes))
        for row in cursor.fetchall():
            key = (row['hsn_code'] or '').strip()
            if key in existing:
                ambiguous.append(key)       # the code is already on file twice
            else:
                existing[key] = row

    # Items named in the file that do not exist
    item_codes = [r['item_code'] for r in rows if r.get('item_code')]
    missing_items = []
    if item_codes:
        ph = ','.join(['%s'] * len(item_codes))
        cursor.execute(f"""SELECT item_code FROM items
                            WHERE item_code IN ({ph}) AND is_deleted = 0""",
                       tuple(item_codes))
        found = {r['item_code'] for r in cursor.fetchall()}
        missing_items = sorted(set(item_codes) - found)

    seen = {}
    for r in rows:
        seen[r['hsn_code']] = seen.get(r['hsn_code'], 0) + 1
    dupes = {k: v for k, v in seen.items() if v > 1}

    # The arithmetic that makes a tax line trustworthy. cgst + sgst is what a
    # customer in this state pays; igst is what one outside it pays. Both
    # should come to the same total, and that total is the rate on the line.
    bad_math, no_rate, updates, creates = [], [], [], []
    for r in rows:
        rate = r['tax_rate']
        if rate is None:
            no_rate.append(r['hsn_code'])
        else:
            split = round((r['cgst'] or 0) + (r['sgst'] or 0), 2)
            if abs(split - rate) > 0.01:
                bad_math.append({'hsn_code': r['hsn_code'], 'why':
                                 f"cgst {r['cgst']} + sgst {r['sgst']} = {split}, "
                                 f"but the rate says {rate}"})
            elif r['igst'] is not None and abs(r['igst'] - rate) > 0.01:
                bad_math.append({'hsn_code': r['hsn_code'], 'why':
                                 f"igst {r['igst']} does not match the rate {rate}"})

        old = existing.get(r['hsn_code'])
        if old:
            changes = []
            for field in ('description', 'tax_rate', 'cgst', 'sgst', 'igst'):
                new_v = r.get(field)
                if new_v is None:
                    continue
                was = old.get(field)
                if str(was) != str(new_v) and _norm_space(str(was or '')) != _norm_space(str(new_v)):
                    changes.append({'field': field, 'was': was if was is not None else '—',
                                    'now': new_v})
            updates.append({'hsn_code': r['hsn_code'], 'changes': changes})
        else:
            creates.append(r['hsn_code'])

    return {
        'total_rows': len(rows),
        'new_codes': creates,
        'existing_codes': updates,
        'duplicates_in_file': dupes,
        'already_duplicated_in_db': sorted(set(ambiguous)),
        'tax_mismatches': bad_math,
        'rows_without_a_rate': no_rate,
        'items_not_found': missing_items,
        'items_named': len(item_codes),
    }


def apply_hsn_import(cursor, rows, default_date_from=None):
    """Write the rows. Matched on hsn_code; blanks never blank a value."""
    default_date_from = default_date_from or date.today()
    created, updated, linked = [], [], 0
    # Only written if the table has somewhere to put it, so the import still
    # runs on a schema without the column.
    keep_source = 'source_id' in {c['COLUMN_NAME']
                                 for c in get_table_columns(cursor, 'hsn')}

    for r in rows:
        code = (r['hsn_code'] or '').strip()
        if not code:
            continue
        cursor.execute("""SELECT * FROM hsn
                           WHERE hsn_code = %s AND COALESCE(is_deleted,0) = 0
                           ORDER BY hsn_id LIMIT 1""", (code,))
        existing = cursor.fetchone()

        if existing:
            hsn_id = existing['hsn_id']
            sets, params = [], []
            fields = ['description', 'tax_rate', 'cgst', 'sgst', 'igst',
                      'tax_date_from', 'tax_date_to']
            if keep_source:
                fields.append('source_id')
            for field in fields:
                value = r.get(field)
                if value is None:
                    continue                      # a blank cell changes nothing
                if str(existing.get(field)) != str(value):
                    sets.append(f"`{field}` = %s")
                    params.append(value)
            if sets:
                params.append(hsn_id)
                cursor.execute(f"UPDATE hsn SET {', '.join(sets)} WHERE hsn_id = %s",
                               tuple(params))
                updated.append(code)
        else:
            # tax_date_from is NOT NULL and the file often omits it.
            names = ['hsn_code', 'description', 'tax_rate', 'cgst', 'sgst',
                     'igst', 'tax_date_from', 'tax_date_to', 'is_deleted']
            vals = [code, r.get('description'), r.get('tax_rate'),
                    r.get('cgst'), r.get('sgst'), r.get('igst'),
                    r.get('tax_date_from') or default_date_from,
                    r.get('tax_date_to'), 0]
            if keep_source:
                names.append('source_id')
                vals.append(r.get('source_id'))
            cursor.execute("INSERT INTO hsn (%s) VALUES (%s)"
                           % (', '.join('`%s`' % n for n in names),
                              ', '.join(['%s'] * len(names))), tuple(vals))
            hsn_id = cursor.lastrowid
            created.append(code)

        # Optional item link. This is how the rest of the application finds an
        # item's tax rate, so a file that names one is doing real work.
        if r.get('item_code'):
            cursor.execute("SELECT item_id FROM items WHERE item_code = %s AND is_deleted = 0 LIMIT 1",
                           (r['item_code'],))
            item = cursor.fetchone()
            if item:
                cursor.execute("""INSERT INTO hsn_items (item_id, hsn_id) VALUES (%s, %s)
                                  ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)""",
                               (item['item_id'], hsn_id))
                linked += 1

    return {'rows': len(rows), 'created': created, 'updated': updated, 'linked': linked}


@app.route('/hsn/import', methods=['GET', 'POST'])
def hsn_import():
    if request.method == 'GET':
        return render_template('hsn_import.html')

    stage = request.form.get('stage', 'preview')

    if stage == 'confirm':
        token = request.form.get('token', '')
        path = os.path.join(tempfile.gettempdir(), f"tribi_hsn_{token}")
        if not token or not os.path.exists(path):
            flash("That upload has expired. Please choose the file again.", "error")
            return redirect(url_for('hsn_import'))
        with open(path, 'rb') as fh:
            file_bytes = fh.read()
        filename = request.form.get('filename', 'upload.xlsx')

        db = get_db()
        cursor = db.cursor(dictionary=True)
        try:
            rows, meta = parse_hsn_file(filename, file_bytes)
            result = apply_hsn_import(cursor, rows)
            db.commit()
            msg = (f"Imported {result['rows']} row(s): "
                   f"{len(result['created'])} HSN code(s) added, "
                   f"{len(result['updated'])} updated.")
            if result['linked']:
                msg += f" {result['linked']} item(s) linked to an HSN code."
            flash(msg, "success")
        except Exception as e:
            db.rollback()
            flash(f"Import failed, nothing was saved: {e}", "error")
        finally:
            cursor.close()
            db.close()
            try:
                os.remove(path)
            except OSError:
                pass
        return redirect(url_for('hsn_list'))

    f = request.files.get('file')
    if not f or not f.filename:
        flash("Choose a file first.", "error")
        return redirect(url_for('hsn_import'))

    file_bytes = f.read()
    db = get_db()
    cursor = db.cursor(dictionary=True)
    try:
        rows, meta = parse_hsn_file(f.filename, file_bytes)
        preview = preview_hsn_import(cursor, rows)
    except Exception as e:
        cursor.close()
        db.close()
        flash(f"Could not read that file: {e}", "error")
        return redirect(url_for('hsn_import'))
    cursor.close()
    db.close()

    token = uuid.uuid4().hex
    with open(os.path.join(tempfile.gettempdir(), f"tribi_hsn_{token}"), 'wb') as fh:
        fh.write(file_bytes)

    return render_template('hsn_import.html', preview=preview, meta=meta,
                           rows=rows[:50], token=token, filename=f.filename)


@app.route('/vendors/import', methods=['GET', 'POST'])
def vendor_import():
    if request.method == 'GET':
        return render_template('vendor_import.html')

    stage = request.form.get('stage', 'preview')

    if stage == 'confirm':
        token = request.form.get('token', '')
        path = os.path.join(tempfile.gettempdir(), f"tribi_vendors_{token}")
        if not token or not os.path.exists(path):
            flash("That upload has expired. Please choose the file again.", "error")
            return redirect(url_for('vendor_import'))
        with open(path, 'rb') as fh:
            file_bytes = fh.read()
        filename = request.form.get('filename', 'upload.xlsx')

        db = get_db()
        cursor = db.cursor(dictionary=True)
        try:
            rows, meta = parse_vendor_file(filename, file_bytes)
            result = apply_vendor_import(cursor, rows)
            db.commit()
            msg = (f"Imported {result['rows']} row(s): "
                   f"{len(result['created'])} vendor(s) added, "
                   f"{len(result['updated'])} updated.")
            if result['revived']:
                msg += (f" {len(result['revived'])} previously deleted vendor(s) "
                        f"restored: {', '.join(result['revived'][:5])}.")
            if result['truncated']:
                bits = ', '.join(f"{t['short_name']}.{t['field']} "
                                 f"({t['was']}->{t['kept']} chars)"
                                 for t in result['truncated'][:5])
                msg += (f" {len(result['truncated'])} value(s) were too long for "
                        f"their column and were cut: {bits}.")
            flash(msg, "success")
        except Exception as e:
            db.rollback()
            flash(f"Import failed, nothing was saved: {e}", "error")
        finally:
            cursor.close()
            db.close()
            try:
                os.remove(path)
            except OSError:
                pass
        return redirect(url_for('vendor_import'))

    f = request.files.get('file')
    if not f or not f.filename:
        flash("Choose a file first.", "error")
        return redirect(url_for('vendor_import'))

    file_bytes = f.read()
    db = get_db()
    cursor = db.cursor(dictionary=True)
    try:
        rows, meta = parse_vendor_file(f.filename, file_bytes)
        preview = preview_vendor_import(cursor, rows)
    except Exception as e:
        cursor.close()
        db.close()
        flash(f"Could not read that file: {e}", "error")
        return redirect(url_for('vendor_import'))
    cursor.close()
    db.close()

    token = uuid.uuid4().hex
    with open(os.path.join(tempfile.gettempdir(), f"tribi_vendors_{token}"), 'wb') as fh:
        fh.write(file_bytes)

    return render_template('vendor_import.html', preview=preview, meta=meta,
                           rows=rows[:50], token=token, filename=f.filename)


@app.route('/storage/import', methods=['GET', 'POST'])
def storage_import():
    """Two steps on purpose: upload and review, then confirm. These files carry
    thousands of rows and create manufacturers as a side effect, which is not
    something to set off with a single click."""
    if request.method == 'GET':
        return render_template('storage_import.html')

    stage = request.form.get('stage', 'preview')
    default_location = (request.form.get('default_location') or 'Main Store').strip()

    if stage == 'confirm':
        token = request.form.get('token', '')
        path = os.path.join(tempfile.gettempdir(), f"tribi_import_{token}")
        if not token or not os.path.exists(path):
            flash("That upload has expired. Please choose the file again.", "error")
            return redirect(url_for('storage_import'))
        with open(path, 'rb') as fh:
            file_bytes = fh.read()
        filename = request.form.get('filename', 'upload.xls')

        db = get_db()
        cursor = db.cursor(dictionary=True)
        try:
            rows, meta = parse_inventory_file(filename, file_bytes)

            # The reviewer ticked the names that really are manufacturers. Work
            # out the rejects by re-deriving the proposed list from the file
            # rather than trusting a hidden field, so a tampered or stale form
            # cannot smuggle a name past the review.
            preview = preview_inventory_import(cursor, rows)
            proposed = {e['name'].strip().lower() for e in preview['new_manufacturers']}
            kept = {v.strip().lower() for v in request.form.getlist('keep_mfr') if v.strip()}
            not_manufacturers = proposed - kept

            result = apply_inventory_import(cursor, rows, default_location,
                                           not_manufacturers=not_manufacturers)
            db.commit()
            msg = (f"Imported {result['rows']} row(s): {result['items_created']} new item(s), "
                   f"{result['items_updated']} item(s) updated, "
                   f"{result['storage_rows_created']} new storage record(s).")
            if result['manufacturers_created']:
                msg += f" {len(result['manufacturers_created'])} manufacturer(s) added."
            if not_manufacturers:
                msg += (f" {len(not_manufacturers)} name(s) kept as part of the MPN "
                        f"instead of becoming manufacturers.")
            if result['units_created']:
                msg += f" Units added: {', '.join(result['units_created'])}."
            if result['desc_changed']:
                msg += f" {len(result['desc_changed'])} description(s) replaced by the file."
            if result['mfr_changed']:
                msg += f" {len(result['mfr_changed'])} manufacturer(s) changed."
            flash(msg, "success")
        except Exception as e:
            db.rollback()
            flash(f"Import failed, nothing was saved: {e}", "error")
            return redirect(url_for('storage_import'))
        finally:
            cursor.close()
            db.close()
            try:
                os.remove(path)
            except OSError:
                pass
        return redirect(url_for('storage_list'))

    # --- preview ---
    f = request.files.get('excel_file')
    if not f or f.filename == '':
        flash("Please choose a file to import.", "error")
        return redirect(url_for('storage_import'))

    file_bytes = f.read()
    try:
        rows, meta = parse_inventory_file(f.filename, file_bytes)
    except Exception as e:
        flash(f"Could not read that file: {e}", "error")
        return redirect(url_for('storage_import'))

    if not rows:
        flash("No item rows were found in that file. Check that it has an item "
              "code column and a stock column.", "error")
        return redirect(url_for('storage_import'))

    db = get_db()
    cursor = db.cursor(dictionary=True)
    try:
        summary = preview_inventory_import(cursor, rows)
    finally:
        cursor.close()
        db.close()

    token = uuid.uuid4().hex
    with open(os.path.join(tempfile.gettempdir(), f"tribi_import_{token}"), 'wb') as fh:
        fh.write(file_bytes)

    return render_template('storage_import.html',
                           preview=summary, meta=meta, token=token,
                           filename=f.filename, default_location=default_location,
                           sample=rows[:15])


# ─── STORAGE LIST ───────────────────────────────────────
# ─── UNIT CONVERSIONS ───────────────────────────────────
# Per-item factors for parts bought in one kind of measure and held in
# another. Conversions within a family (kg to g, m to cm) live in `units` and
# are not maintained here.

@app.route('/unit-conversions')
def unit_conversion_list():
    db = get_db()
    cursor = db.cursor(dictionary=True)

    cursor.execute("""SELECT c.*, i.item_code, i.`desc`, i.unit_id AS item_unit_id,
                             au.unit_short_name AS alt_unit,
                             su.unit_short_name AS stock_unit,
                             cu.unit_short_name AS current_unit
                        FROM item_unit_conversion c
                        JOIN items i   ON i.item_id   = c.item_id
                        JOIN units au  ON au.unit_id  = c.unit_id
                        JOIN units su  ON su.unit_id  = c.stock_unit_id
                        LEFT JOIN units cu ON cu.unit_id = i.unit_id
                       ORDER BY i.item_code, au.unit_short_name""")
    rows = cursor.fetchall()

    stale = []
    for r in rows:
        num, den = float(r['conv_fact_num'] or 0), float(r['conv_fact_denom'] or 1)
        # Both directions, because only the second is checkable by a person.
        r['per_alt'] = round(num / den, 6) if den else 0
        r['per_stock'] = round(den / num, 4) if num else 0
        # A factor measured against metres means nothing if the item now holds
        # centimetres. Flag it rather than quietly converting by the wrong amount.
        r['is_stale'] = (r['item_unit_id'] is None
                         or int(r['stock_unit_id']) != int(r['item_unit_id']))
        if r['is_stale']:
            stale.append({'item_code': r['item_code'],
                          'measured_against': r['stock_unit'],
                          'now_holds': r['current_unit']})

    cursor.execute("""SELECT unit_id, unit_name, unit_short_name
                        FROM units ORDER BY unit_name""")
    units = cursor.fetchall()

    cursor.execute("""SELECT i.item_id, i.item_code, i.`desc` AS d,
                             u.unit_short_name AS unit
                        FROM items i
                        LEFT JOIN units u ON u.unit_id = i.unit_id
                       WHERE i.is_deleted = 0
                       ORDER BY i.item_code""")
    items = [{'item_id': r['item_id'], 'code': r['item_code'],
              'desc': (r['d'] or '')[:60], 'unit': r['unit']}
             for r in cursor.fetchall()]

    cursor.close()
    db.close()
    return render_template('unit_conversion.html', rows=rows, units=units,
                           stale=stale, items_json=json.dumps(items))


@app.route('/unit-conversions/add', methods=['POST'])
def unit_conversion_add():
    item_id = request.form.get('item_id')
    unit_id = request.form.get('unit_id')
    remarks = (request.form.get('remarks') or '').strip()[:100]
    try:
        num = int(request.form.get('conv_fact_num') or 0)
        den = int(request.form.get('conv_fact_denom') or 0)
    except ValueError:
        num = den = 0

    # A zero or negative factor would collapse every quantity it touched, and
    # unlike a missing row it would do so silently.
    if not item_id or not unit_id or num <= 0 or den <= 0:
        flash('Pick an item and a unit, and give both numbers as whole numbers above zero.', 'error')
        return redirect(url_for('unit_conversion_list'))

    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT unit_id FROM items WHERE item_id = %s", (item_id,))
    item = cursor.fetchone()
    if not item or not item['unit_id']:
        cursor.close()
        db.close()
        flash('That item has no stock unit, so there is nothing to convert into. '
              'Set its unit first.', 'error')
        return redirect(url_for('unit_conversion_list'))
    if int(item['unit_id']) == int(unit_id):
        cursor.close()
        db.close()
        flash('That is the unit the item is already held in — nothing to convert.', 'error')
        return redirect(url_for('unit_conversion_list'))

    try:
        cursor.execute("""INSERT INTO item_unit_conversion
                              (item_id, unit_id, stock_unit_id, conv_fact_num,
                               conv_fact_denom, remarks, updated_by, updated_at)
                          VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())
                          ON DUPLICATE KEY UPDATE
                              stock_unit_id   = VALUES(stock_unit_id),
                              conv_fact_num   = VALUES(conv_fact_num),
                              conv_fact_denom = VALUES(conv_fact_denom),
                              remarks         = VALUES(remarks),
                              updated_by      = VALUES(updated_by),
                              updated_at      = NOW()""",
                       (item_id, unit_id, item['unit_id'], num, den, remarks or None,
                        session.get('username') or 'unknown'))
        db.commit()
        flash('Conversion saved.', 'success')
    except Exception as e:
        db.rollback()
        flash(f'Could not save that: {e}', 'error')
    cursor.close()
    db.close()
    return redirect(url_for('unit_conversion_list'))


@app.route('/unit-conversions/<int:conv_id>/delete', methods=['POST'])
def unit_conversion_delete(conv_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    try:
        cursor.execute("DELETE FROM item_unit_conversion WHERE conv_id = %s", (conv_id,))
        db.commit()
        # Removing a row is safe: without it receipts and issues still go
        # through, they simply lose the check against the order.
        flash('Conversion removed.', 'success')
    except Exception as e:
        db.rollback()
        flash(f'Could not remove that: {e}', 'error')
    cursor.close()
    db.close()
    return redirect(url_for('unit_conversion_list'))


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

    # ─── AVAILABILITY, IN ONE PASS ──────────────────────────────
    #
    # This used to call get_item_stock() once per row. That helper runs three
    # queries, so a store of 4,300 items meant about thirteen thousand round
    # trips to the database before the page could start rendering — which is
    # why this screen felt slower than every other list.
    #
    # The same two figures are now fetched as two aggregates over the whole
    # table. The arithmetic is unchanged: available = what is held, minus what
    # is actively blocked, floored at zero.
    blocked_by_item, held_by_item = {}, {}
    try:
        cursor.execute("""SELECT item_id, COALESCE(SUM(qty_blocked), 0) AS blocked
                            FROM stock_blocking
                           WHERE block_status = 0
                           GROUP BY item_id""")
        blocked_by_item = {int(r['item_id']): float(r['blocked'] or 0)
                           for r in cursor.fetchall()}
    except Exception:
        blocked_by_item = {}      # table absent; same fallback the helper had

    # Summed per item rather than read off the row, because get_item_stock
    # totalled every storage row for an item. There is one row per item today,
    # but matching the old behaviour exactly costs nothing.
    cursor.execute("""SELECT item_id, COALESCE(SUM(physical_availability), 0) AS held
                        FROM storage GROUP BY item_id""")
    held_by_item = {int(r['item_id']): float(r['held'] or 0)
                    for r in cursor.fetchall()}

    total_physical_availability = 0.0
    for s in storages:
        iid = int(s['item_id'])
        held = held_by_item.get(iid, float(s.get('physical_availability') or 0))
        s['currently_available'] = max(0.0, round(held - blocked_by_item.get(iid, 0.0),
                                                  QTY_DP))
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

        # The HSN lives on the item, in hsn_items — storage has no column for
        # it and one row per item, so classifying here classifies the item
        # everywhere. An item that is already classified keeps its code;
        # reclassifying is a decision for the HSN screen, not a side
        # effect of entering stock.
        #
        # Offered, not demanded. Receiving stock is not a tax document, so
        # a missing HSN must not stop someone recording what is on the
        # shelf. The purchase order is where it is insisted on, because
        # that is where the rate is actually printed.
        resolved_hsn = set_item_hsn(cursor, item_id, request.form.get('hsn_id'))

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

        if not resolved_hsn:
            flash('This item has no HSN code yet. A purchase order will ask '
                  'for one before it can show tax.', 'warning')

        db.commit()
        cursor.close()
        db.close()
        return redirect(url_for('storage_list'))

    cursor.execute("SELECT item_id, item_code, `desc` as item_desc FROM items WHERE is_deleted=0 ORDER BY item_code")
    items = cursor.fetchall()
    hsns = hsn_choices(cursor)
    cursor.close()
    db.close()
    return render_template('storage_add.html', items=items, hsns=hsns)


# ─── API ITEM STORAGE LOCATION ─────────────────────────
@app.route('/api/item-storage-location/<int:item_id>')
def api_item_storage_location(item_id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT store_location, physical_availability FROM storage WHERE item_id = %s LIMIT 1", (item_id,))
    row = cursor.fetchone()
    # Whether this item is classified decides if the form shows its HSN or
    # asks for one, so the same call that fills the location answers that too.
    hsn = hsn_for_item(cursor, item_id)
    cursor.close()
    db.close()
    return jsonify({
        'exists': bool(row),
        'location': (row['store_location'] if row else None),
        'physical_availability': (float(row['physical_availability'] or 0) if row else None),
        'hsn_id': (hsn['hsn_id'] if hsn else None),
        'hsn_code': (hsn['hsn_code'] if hsn else None),
        'hsn_rate': (float(hsn['tax_rate']) if hsn and hsn['tax_rate'] is not None else None)
    })

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
        item_id = form_item_id(request.form)
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

            # Linked only if an item was chosen. A code with nothing under
            # it is perfectly ordinary — items are attached later, by the
            # import, by the storage screen, or on a purchase order line.
            if item_id:
                cursor.execute("""
                    INSERT INTO hsn_items (item_id, hsn_id)
                    VALUES (%s, %s)
                    ON DUPLICATE KEY UPDATE hsn_id = VALUES(hsn_id)
                """, (item_id, hsn_id))
            db.commit()
            flash(f"HSN Code {hsn_code} created" +
                  (" and linked to the selected item." if item_id
                   else ". No item is classified under it yet."), "success")
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
        item_id = form_item_id(request.form)
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

            # Leaving the item blank means "do not change the link", not
            # "remove it". Clearing a classification is a deliberate act and
            # should not happen because a field was left empty.
            if item_id and old_item_id is not None and old_item_id != item_id:
                cursor.execute("DELETE FROM hsn_items WHERE item_id = %s", (old_item_id,))
            
            # Migrated from hsn.item_id — now sourced via hsn_items junction table
            cursor.execute("""
                UPDATE hsn
                SET hsn_code = %s, description = %s, tax_rate = %s, cgst = %s, sgst = %s, igst = %s, tax_date_from = %s
                WHERE hsn_id = %s
            """, (hsn_code, description, tax_rate, cgst, sgst, igst, tax_date_from, hsn_id))
            
            # Migrated from items.hsn_id — now sourced via hsn_items junction table
            if item_id:
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

        # Same availability guard the work order Issue button uses: free stock,
        # plus anything already blocked for this WO (that reservation is ours to draw on).
        stock = get_item_stock(cursor, item_id)
        wo_blocked = 0.0
        if wo_id:
            cursor.execute("""
                SELECT COALESCE(SUM(qty_blocked), 0) as wo_blocked
                FROM stock_blocking WHERE item_id=%s AND wo_id=%s AND block_status=0
            """, (item_id, wo_id))
            wo_blocked = float(cursor.fetchone()['wo_blocked'] or 0)
        real_available = round(stock['available'] + wo_blocked, QTY_DP)

        if qty_issued > real_available:
            cursor.close()
            db.close()
            flash(f"Cannot issue {qty_issued} of '{item_code}' — only {real_available} available.", 'error')
            return redirect(url_for('issue_transaction_list'))

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

        # Release blocks and reduce physical stock — the step this form was missing.
        if not apply_issue_to_stock(cursor, item_id, wo_id, qty_issued):
            # Same race as the work order Issue button. The transaction row and
            # the work_order_item update above are inside this rollback too, so
            # nothing survives a failed deduction.
            db.rollback()
            cursor.close()
            db.close()
            flash(f"Could not issue {_trim_number(qty_issued)} of '{item_code}' — "
                  f"the stock was taken between checking and issuing. Nothing "
                  f"was changed.", 'error')
            return redirect(url_for('issue_transaction_list'))

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
        conn = mysql.connector.connect(**DB_CONFIG)
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


def ensure_schema_columns():
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True)
        
        # HSN codes and the item -> HSN links used to be generated here:
        # eight hardcoded codes were inserted on every start, and every
        # item without a link was given one by keyword-matching its
        # description, with 854370 as the catch-all. That produced 4,322
        # links of which 1,472 were the fallback, and misfiled anything
        # whose description happened to contain 'res', 'cap', 'plate' or
        # 'filter'. HSN now comes from the import file only, so the
        # guessing is gone. An item with no HSN shows no tax rate, which
        # is the honest answer and is visible, where a wrong rate is not.

        # Add is_locked column if not exists
        cursor.execute("SHOW COLUMNS FROM purchase_order LIKE 'is_locked'")
        if not cursor.fetchone():
            cursor.execute("ALTER TABLE purchase_order ADD COLUMN is_locked TINYINT(1) NOT NULL DEFAULT 0")
            conn.commit()
            print("Added is_locked column to purchase_order table.")

        # The unique index on hsn.hsn_code used to be dropped here on
        # every start, because the seeding above could insert a code the
        # table already held. Nothing generates codes any more, so there
        # is nothing to make room for, and dropping an index the database
        # is relied on to hold would only undo it again next start.

        # Add currency column to purchase_order if not exists
        cursor.execute("SHOW COLUMNS FROM purchase_order LIKE 'currency'")
        if not cursor.fetchone():
            cursor.execute("ALTER TABLE purchase_order ADD COLUMN currency VARCHAR(10) NOT NULL DEFAULT 'INR'")
            conn.commit()
            print("Added currency column to purchase_order table.")

        cursor.close()
        conn.close()
    except Exception as e:
        print("Schema check error:", e)


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
        'app_version': APP_VERSION,
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
    ensure_schema_columns()

    # ─── DEBUG IS OFF UNLESS ASKED FOR ──────────────────────────
    #
    # debug=True on host 0.0.0.0 hands a full traceback — file paths, source
    # lines, local variables — to anybody on the network who can make the
    # application throw. use_evalex=False already blocked the interactive
    # console, which was the worst of it, but the disclosure remains.
    #
    # Set TRIBI_DEBUG=1 to turn it back on while developing.
    debug_mode = os.environ.get('TRIBI_DEBUG', '').strip().lower() in ('1', 'true', 'yes')

    if not debug_mode:
        # With debug off an unhandled error becomes a blank 500 page, so the
        # traceback has to go somewhere a person can read it. Without this,
        # diagnosing a fault in the store means asking someone to reproduce it
        # while you watch.
        import logging
        from logging.handlers import RotatingFileHandler
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                'tribi_errors.log')
        handler = RotatingFileHandler(log_path, maxBytes=2_000_000, backupCount=5)
        handler.setLevel(logging.ERROR)
        handler.setFormatter(logging.Formatter(
            '%(asctime)s  %(levelname)s  %(message)s\n'
            '  %(pathname)s:%(lineno)d\n'))
        app.logger.addHandler(handler)
        app.logger.setLevel(logging.ERROR)
        print(f"Errors will be written to {log_path}")

    print(f"Starting Tribi ERP — debug {'ON' if debug_mode else 'off'}")
    app.run(host='0.0.0.0', port=5000, debug=debug_mode, use_evalex=False)

#