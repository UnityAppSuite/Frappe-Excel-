# Copyright (c) 2026, Ujjwal Aggrawal and contributors
# For license information, please see license.txt

"""
excel_view.api — server-side endpoints for Excel View.

All methods are @frappe.whitelist(), meaning they are accessible via
frappe.call({ method: "excel_view.api.<name>", ... }) from the client.

Security model:
  - Any logged-in user can create workbooks.
  - A workbook is readable by its owner or by anyone if is_public = 1.
  - Only the owner (or System Manager) can update or delete a workbook.
  - Every method verifies that the caller has READ permission on the target
    DocType — so users can't save views for doctypes they can't access.

V2.3 — Frappe Formula Library
  Scalar async formula functions callable from HyperFormula cells:
    FRAPPE_GET(doctype, name, fieldname)
    FRAPPE_SUM(doctype, fieldname [, fk, fv …])
    FRAPPE_COUNT(doctype [, fk, fv …])
    FRAPPE_AVG(doctype, fieldname [, fk, fv …])
    GL_BALANCE(account, company [, from_date, to_date, cost_center, finance_book])
    STOCK_QTY(item_code, warehouse [, as_of_date])
    ITEM_PRICE(item_code, price_list [, qty, customer, uom])
"""

import frappe
from frappe import _

from excel_view.decorators import excel_whitelist

# orjson: 5-10x faster JSON parsing/serialization than stdlib json.
# Used for all our heavy payload parsing (bulk_fetch, smart_lookup, etc.)
try:
	import orjson as _json_lib

	def _fast_parse(s):
		if isinstance(s, bytes | bytearray):
			return _json_lib.loads(s)
		return _json_lib.loads(s.encode() if isinstance(s, str) else s)

	def _fast_dumps(obj):
		return _json_lib.dumps(obj).decode()
except ImportError:
	import json as _json_lib

	_fast_parse = _json_lib.loads
	_fast_dumps = _json_lib.dumps


def _check_doctype_permission(doctype, ptype="read"):
	"""Check permission, skipping child table DocTypes (they inherit from parent)."""
	meta = frappe.get_meta(doctype)
	if meta.istable:
		return  # Child tables inherit permissions from parent; skip standalone check
	frappe.has_permission(doctype, ptype, throw=True)


# ── V2.3 helpers ──────────────────────────────────────────────────────────────

#: Fields that exist on every DocType but are NOT in meta.fields — always valid.
_SYSTEM_FIELDS: frozenset[str] = frozenset(
	{
		"name",
		"owner",
		"creation",
		"modified",
		"modified_by",
		"docstatus",
		"idx",
		"_user_tags",
		"_assign",
		"_comments",
		"_seen",
	}
)

#: Hard cap for frappe_aggregate SUM/AVG to avoid runaway DB scans.
_AGGREGATE_ROW_CAP = 50_000


def _validate_fieldname(doctype: str, fieldname: str) -> None:
	"""
	Raise DoesNotExistError if *fieldname* is not a valid field on *doctype*.

	Checks system fields first (always allowed), then the live meta — which
	includes custom fields added by any installed app.  This makes the
	validation fully dynamic without any hardcoded field lists.
	"""
	if fieldname in _SYSTEM_FIELDS:
		return

	meta = frappe.get_meta(doctype)
	valid_fields = {df.fieldname for df in meta.fields}

	if fieldname not in valid_fields:
		frappe.throw(
			_("Field '{0}' does not exist on DocType '{1}'.").format(fieldname, doctype),
			frappe.DoesNotExistError,
		)


_FILTER_OPS = frozenset({">", "<", ">=", "<=", "!=", "like", "not like", "in", "not in", "between"})


def _parse_filter_pairs(
	fk1=None,
	fv1=None,
	fk2=None,
	fv2=None,
	fk3=None,
	fv3=None,
	fk4=None,
	fv4=None,
) -> list:
	"""
	Build a Frappe filter list from up to four SUMIF-style key/value pairs.

	Supports both exact-match and operator-style keys:
	  Exact:    fk="ev_sales_person", fv="Arjun" → ["ev_sales_person", "=", "Arjun"]
	  Operator: fk="transaction_date >=", fv="2026-03-01" → ["transaction_date", ">=", "2026-03-01"]

	Keys/values that are None or empty are silently skipped.
	Returns a list of [field, op, value] triples (frappe.get_list accepts this format).
	"""
	filters = []
	for key, val in ((fk1, fv1), (fk2, fv2), (fk3, fv3), (fk4, fv4)):
		if not key or val is None or val == "":
			continue
		key = str(key).strip()
		# Detect trailing operator: "transaction_date >=" → field="transaction_date", op=">="
		parts = key.rsplit(" ", 1)
		if len(parts) == 2 and parts[1].lower() in _FILTER_OPS:
			filters.append([parts[0], parts[1], val])
		else:
			filters.append([key, "=", val])
	return filters


# ── Child Table — Inline Expand (V3.5) ───────────────────────────────────────
#
# Error-handling philosophy:
#   Endpoints return {"_ev_error": True, "message": "...", "exc_type": "..."} on
#   failure instead of raising.  This bypasses Frappe's request-level error handler
#   so _server_messages never reach the browser — all error display is handled by
#   the JS EVModal/toast system.


def _make_ev_error(exc: Exception) -> dict:
	"""
	Serialise *exc* into the structured error dict that the front-end
	_ev_call() helper recognises and routes through EVModal.

	Priority for the user-facing message:
	  1. frappe.message_log  — populated by frappe.throw(); contains the
	     translated, HTML-formatted message that Frappe normally shows in its
	     own dialog.  We strip the HTML and use the plain text.
	  2. str(exc)            — raw Python exception message; acceptable but
	     may contain internal identifiers (e.g. MandatoryError format).

	IMPORTANT: call this *before* frappe.message_log.clear() so the log is
	still populated.
	"""
	exc_type = type(exc).__name__.split(".")[-1]

	msg: str | None = None

	# 1 — user-friendly message from frappe.message_log
	if frappe.message_log:
		try:
			entry = frappe.message_log[-1]
			if isinstance(entry, dict):
				msg = entry.get("message") or entry.get("title")
			else:
				parsed = (_fast_parse(entry) if isinstance(entry, str) else (entry or [])) or {}
				msg = parsed.get("message") or parsed.get("title")
		except Exception:
			pass

	# 2 — fallback to exception string
	if not msg:
		msg = str(exc) if exc else frappe._("An unexpected error occurred.")

	return {
		"_ev_error": True,
		"message": frappe.utils.strip_html(msg),
		"exc_type": exc_type,
	}


#
# Four endpoints that power the "expand parent row → inline child grid" feature.
# All writes go through the parent doc.save() so every controller, validate,
# before_save and after_save hook runs exactly as it would from the Frappe form.
#
# Security model:
#   - READ  endpoints: caller must have READ on the parent DocType.
#   - WRITE endpoints: caller must have WRITE on the parent DocType.
#   - Child tables (istable=1) inherit perms from their parent; no separate
#     permission row is required on the child DocType itself.
#   - All fieldnames are validated against live meta — no raw SQL injection.

#: Field types that carry no storable value (sections, columns, HTML, etc.)
_NO_VALUE_TYPES: frozenset[str] = frozenset(
	{
		"Section Break",
		"Column Break",
		"Tab Break",
		"HTML",
		"Fold",
		"Heading",
		"Button",
		"Image",
		"Barcode",
		"Geolocation",
		"Table",
		"Table MultiSelect",
	}
)


def _get_child_table_df(parent_meta, child_fieldname: str):
	"""
	Return the field descriptor for *child_fieldname* on *parent_meta*.
	Raises ValidationError if the field does not exist or is not a Table type.
	"""
	df = next(
		(
			f
			for f in parent_meta.fields
			if f.fieldname == child_fieldname and f.fieldtype in ("Table", "Table MultiSelect")
		),
		None,
	)
	if not df:
		frappe.throw(
			_("'{0}' is not a valid child table field on {1}.").format(child_fieldname, parent_meta.name)
		)
	return df


def _child_col_defs(child_meta) -> list[dict]:
	"""
	Return a JSON-serialisable list of column definitions for the child DocType.
	Only includes fields with a storable value (excludes Section Break etc.).
	The 'name' system field is prepended so each row has a stable identifier.
	`allow_on_submit` mirrors Frappe's perm.js logic: tells the UI whether this
	field may be edited on an already-submitted parent document.
	"""
	cols = []
	for f in child_meta.fields:
		if f.fieldtype in _NO_VALUE_TYPES:
			continue
		cols.append(
			{
				"fieldname": f.fieldname,
				"label": _(f.label or f.fieldname),
				"fieldtype": f.fieldtype,
				"options": f.options or "",
				"reqd": bool(f.reqd),
				"read_only": bool(f.read_only),
				"hidden": bool(f.hidden),
				"in_list_view": bool(f.in_list_view),
				"default": f.default or "",
				"allow_on_submit": bool(f.allow_on_submit),  # V3.5: submitted-doc edit guard
			}
		)
	return cols


@frappe.whitelist()
def get_child_rows(doctype: str, parent_name: str, child_fieldname: str) -> dict:
	"""
	Return all rows of one child table for a single parent document.

	Args:
	    doctype:         Parent DocType (e.g. "Sales Order")
	    parent_name:     Name of the parent document (e.g. "SAL-ORD-2026-00001")
	    child_fieldname: Fieldname of the child table on the parent (e.g. "items")

	Returns:
	    {
	      "child_doctype": "Sales Order Item",
	      "fields": [ { fieldname, label, fieldtype, options, reqd, ... }, ... ],
	      "rows":   [ { name, idx, fieldname: value, ... }, ... ]
	    }

	Security: READ on parent DocType is required.
	"""
	frappe.has_permission(doctype, "read", parent_name, throw=True)

	parent_meta = frappe.get_meta(doctype)
	df = _get_child_table_df(parent_meta, child_fieldname)
	child_doctype: str = df.options

	# V3.5 — Submitted-doc edit guard: fetch parent docstatus in one cheap query
	parent_docstatus: int = frappe.db.get_value(doctype, parent_name, "docstatus") or 0

	child_meta = frappe.get_meta(child_doctype)
	col_defs = _child_col_defs(child_meta)

	# Build the field list for the DB query — system fields + all value-bearing fields
	fetch_fields = ["name", "idx"] + [c["fieldname"] for c in col_defs]
	# Deduplicate while preserving order (name/idx might appear in col_defs too)
	seen: set[str] = set()
	safe_fields: list[str] = []
	for f in fetch_fields:
		if f not in seen:
			seen.add(f)
			safe_fields.append(f)

	rows = frappe.get_all(
		child_doctype,
		filters={
			"parent": parent_name,
			"parentfield": child_fieldname,
			"parenttype": doctype,
		},
		fields=safe_fields,
		order_by="idx asc",
		limit=0,
	)

	return {
		"child_doctype": child_doctype,
		"fields": col_defs,
		"rows": [dict(r) for r in rows],
		# V3.5 — submitted-doc awareness
		"parent_docstatus": parent_docstatus,
		"table_allow_on_submit": bool(df.allow_on_submit),  # controls add/delete
	}


@frappe.whitelist()
def save_child_row(
	doctype: str,
	parent_name: str,
	child_fieldname: str,
	child_name: str,
	fields: str,
) -> dict:
	"""
	Update one child row via the full Frappe ORM (parent_doc.save()).

	Runs all controller hooks (validate, before_save, on_change, after_save)
	— same as saving from the Frappe form UI.

	Security: WRITE on parent DocType is required.
	"""
	try:
		frappe.has_permission(doctype, "write", parent_name, throw=True)

		fields_dict: dict = _fast_parse(fields) if isinstance(fields, str) else fields
		if not fields_dict:
			return {"name": child_name}

		parent_meta = frappe.get_meta(doctype)
		df = _get_child_table_df(parent_meta, child_fieldname)
		child_doctype: str = df.options

		child_meta = frappe.get_meta(child_doctype)
		valid_fields = {f.fieldname for f in child_meta.fields} | _SYSTEM_FIELDS
		for fn in fields_dict:
			if fn not in valid_fields:
				frappe.throw(_("Field '{0}' does not exist on {1}.").format(fn, child_doctype))

		parent_doc = frappe.get_doc(doctype, parent_name)
		target_child = next(
			(c for c in (parent_doc.get(child_fieldname) or []) if c.name == child_name),
			None,
		)
		if not target_child:
			frappe.throw(
				_("Child row '{0}' not found in {1}.{2}").format(child_name, parent_name, child_fieldname)
			)

		for fn, val in fields_dict.items():
			target_child.set(fn, val)

		parent_doc.save()
		frappe.db.commit()
		return {"name": child_name}

	except Exception as exc:
		ev_err = _make_ev_error(exc)
		frappe.message_log.clear()
		frappe.db.rollback()
		return ev_err


@frappe.whitelist()
def insert_child_row(
	doctype: str,
	parent_name: str,
	child_fieldname: str,
	fields: str,
) -> dict:
	"""
	Append a new child row via the full Frappe ORM (parent_doc.append + save).

	Runs all controller hooks — identical to adding a row from the Frappe form UI.

	Returns: { "name": new_child_name, "idx": new_idx }
	Security: WRITE on parent DocType is required.
	"""
	try:
		frappe.has_permission(doctype, "write", parent_name, throw=True)

		fields_dict: dict = _fast_parse(fields) if isinstance(fields, str) else (fields or {})

		parent_meta = frappe.get_meta(doctype)
		df = _get_child_table_df(parent_meta, child_fieldname)
		child_doctype: str = df.options

		if fields_dict:
			child_meta = frappe.get_meta(child_doctype)
			valid_fields = {f.fieldname for f in child_meta.fields} | _SYSTEM_FIELDS
			for fn in fields_dict:
				if fn not in valid_fields:
					frappe.throw(_("Field '{0}' does not exist on {1}.").format(fn, child_doctype))

		parent_doc = frappe.get_doc(doctype, parent_name)
		child_row = parent_doc.append(child_fieldname, fields_dict)
		parent_doc.save()
		frappe.db.commit()
		return {"name": child_row.name, "idx": child_row.idx}

	except Exception as exc:
		ev_err = _make_ev_error(exc)
		frappe.message_log.clear()
		frappe.db.rollback()
		return ev_err


@frappe.whitelist()
def delete_child_row(
	doctype: str,
	parent_name: str,
	child_fieldname: str,
	child_name: str,
) -> dict:
	"""
	Remove a child row from a parent document via the full Frappe ORM.

	The parent document is loaded, the matching child row is removed from the
	in-memory list, and parent_doc.save() is called so all hooks run correctly
	(e.g. total recalculation, on_change, after_save).

	Args:
	    doctype:         Parent DocType
	    parent_name:     Name of the parent document
	    child_fieldname: Child table fieldname on the parent
	    child_name:      `name` of the child row to delete

	Returns:
	    { "deleted": child_name }

	Security: WRITE on parent DocType is required.
	"""
	try:
		frappe.has_permission(doctype, "write", parent_name, throw=True)

		parent_meta = frappe.get_meta(doctype)
		_get_child_table_df(parent_meta, child_fieldname)

		parent_doc = frappe.get_doc(doctype, parent_name)
		original_children = parent_doc.get(child_fieldname) or []
		filtered = [c for c in original_children if c.name != child_name]

		if len(filtered) == len(original_children):
			frappe.throw(
				_("Child row '{0}' not found in {1}.{2}").format(child_name, parent_name, child_fieldname)
			)

		parent_doc.set(child_fieldname, filtered)
		parent_doc.save()
		frappe.db.commit()
		return {"deleted": child_name}

	except Exception as exc:
		ev_err = _make_ev_error(exc)  # capture friendly msg BEFORE clearing log
		frappe.message_log.clear()
		frappe.db.rollback()
		return ev_err


@frappe.whitelist()
def save_child_table_bulk(
	doctype: str,
	parent_name: str,
	child_fieldname: str,
	changes: str,
) -> dict:
	"""
	Save ALL pending child-table changes in a single transaction.

	``changes`` is a JSON object::

	    {
	      "updates": [{"name": "<child_name>", "fields": {fn: val, ...}}, ...],
	      "inserts": [{"temp_name": "_new_ct_<ts>", "fields": {fn: val, ...}}, ...]
	    }

	One ``parent_doc.save()`` → one DB transaction → all ORM hooks run once.
	Dramatically faster than N individual save calls (no N round-trips).

	Returns::

	    {"ok": true, "inserted": [{"temp_name": "...", "name": "...", "idx": N}, ...]}

	Security: WRITE on parent DocType required.
	"""
	try:
		frappe.has_permission(doctype, "write", parent_name, throw=True)

		changes_dict: dict = _fast_parse(changes) if isinstance(changes, str) else (changes or {})
		updates: list = changes_dict.get("updates") or []
		inserts: list = changes_dict.get("inserts") or []
		deletes: list = changes_dict.get("deletes") or []

		if not updates and not inserts and not deletes:
			return {"ok": True, "inserted": []}

		parent_meta = frappe.get_meta(doctype)
		df = _get_child_table_df(parent_meta, child_fieldname)
		child_doctype = df.options
		child_meta = frappe.get_meta(child_doctype)
		valid_fields = {f.fieldname for f in child_meta.fields} | _SYSTEM_FIELDS

		# Validate all field names up-front — fail fast before touching the DB
		for upd in updates:
			for fn in upd.get("fields") or {}:
				if fn not in valid_fields:
					frappe.throw(_("Field '{0}' does not exist on {1}.").format(fn, child_doctype))
		for ins in inserts:
			for fn in ins.get("fields") or {}:
				if fn not in valid_fields:
					frappe.throw(_("Field '{0}' does not exist on {1}.").format(fn, child_doctype))

		parent_doc = frappe.get_doc(doctype, parent_name)
		existing = {c.name: c for c in (parent_doc.get(child_fieldname) or [])}

		# Apply deletes — remove from the in-memory child list before save
		# so the controller sees the final desired state (e.g. total % = 100)
		delete_set = set(deletes)
		if delete_set:
			parent_doc.set(
				child_fieldname,
				[row for row in (parent_doc.get(child_fieldname) or []) if row.name not in delete_set],
			)
			existing = {c.name: c for c in (parent_doc.get(child_fieldname) or [])}

		# Apply updates
		for upd in updates:
			child_name = upd.get("name")
			fields_dict = upd.get("fields") or {}
			if not child_name or not fields_dict:
				continue
			if child_name not in existing:
				frappe.throw(_("Child row '{0}' not found.").format(child_name))
			for fn, val in fields_dict.items():
				existing[child_name].set(fn, val)

		# Apply inserts
		new_rows: list = []
		for ins in inserts:
			fields_dict = ins.get("fields") or {}
			new_row = parent_doc.append(child_fieldname, fields_dict)
			new_rows.append((ins.get("temp_name", ""), new_row))

		parent_doc.save()

		# For submitted docs (docstatus=1), Frappe's update_after_submit flow runs
		# on_update_after_submit() AFTER db_update(). Any fields recalculated there
		# (e.g. allocated_amount via calculate_contribution()) are updated in-memory
		# but never flushed to DB. Force-persist them now.
		if parent_doc.docstatus == 1:
			for row in parent_doc.get(child_fieldname) or []:
				row.db_update()

		frappe.db.commit()

		return {
			"ok": True,
			"inserted": [
				{"temp_name": temp_name, "name": row.name, "idx": row.idx} for temp_name, row in new_rows
			],
		}

	except Exception as exc:
		ev_err = _make_ev_error(exc)
		frappe.message_log.clear()
		frappe.db.rollback()
		return ev_err


# ── Child Table Data ─────────────────────────────────────────────────────────


@frappe.whitelist()
def get_child_data(doctype: str, requests: str, parent_names: str) -> dict:
	"""
	Fetch child table fields for a list of parent doc names and aggregate values
	as comma-joined strings.

	Args:
	    requests: JSON dict  {table_fieldname: [child_fieldname, ...]}
	    parent_names: JSON list of parent doc names

	Returns:
	    {parent_name: {"table_fn__child_fn": "val1, val2", ...}}

	Security: requires READ on the parent DocType and on each child DocType.
	Child field names are validated against live meta (SQL-injection safe).
	"""
	frappe.has_permission(doctype, "read", throw=True)

	reqs: dict = (_fast_parse(requests) if isinstance(requests, str) else (requests or [])) or {}
	names: list = (_fast_parse(parent_names) if isinstance(parent_names, str) else (parent_names or [])) or []

	if not reqs or not names:
		return {}

	parent_meta = frappe.get_meta(doctype)
	result: dict = {}

	for table_fn, child_fields in reqs.items():
		# Validate table fieldname exists on parent and is a Table type
		table_df = next(
			(
				df
				for df in parent_meta.fields
				if df.fieldname == table_fn and df.fieldtype in ("Table", "Table MultiSelect")
			),
			None,
		)
		if not table_df:
			continue

		child_doctype: str = table_df.options

		# Child tables (istable=1) inherit perms from parent; _check_doctype_permission
		# skips the check for them, so this is safe without a separate perm row.
		_check_doctype_permission(child_doctype, "read")

		# Validate requested child fields against live meta (prevents SQL injection)
		child_meta = frappe.get_meta(child_doctype)
		valid_child_fields = {df.fieldname for df in child_meta.fields}
		safe_fields = [f for f in child_fields if f in valid_child_fields]
		if not safe_fields:
			continue

		# Fetch all matching child rows in one query
		children = frappe.get_all(
			child_doctype,
			filters={"parent": ["in", names], "parentfield": table_fn},
			fields=["parent", *safe_fields],
			order_by="idx asc",
			ignore_permissions=False,
			limit=0,
		)

		# Group by parent, collect values per field
		parent_buckets: dict = {}
		for child in children:
			p = child.parent
			if p not in parent_buckets:
				parent_buckets[p] = {f: [] for f in safe_fields}
			for f in safe_fields:
				val = child.get(f)
				if val is not None and val != "":
					parent_buckets[p][f].append(str(val))

		# Build composite keys and join values
		for p, fields_data in parent_buckets.items():
			if p not in result:
				result[p] = {}
			for f, values in fields_data.items():
				result[p][f"{table_fn}__{f}"] = ", ".join(values)

	return result


# ── V2.3 formula endpoints ────────────────────────────────────────────────────


@frappe.whitelist()
def frappe_get(doctype: str, name: str, fieldname: str) -> dict:
	"""
	Fetch a single field value from one document.

	Three access patterns are supported:

	1. Plain field:
	       FRAPPE_GET("Customer", "Tata Motors", "credit_limit")

	2. Link-field traversal (dot notation — walks Link hops):
	       FRAPPE_GET("Sales Person", "Arjun Sharma", "employee.ctc")
	       → reads Sales Person.employee (Link→Employee), then Employee.ctc

	3. Child-table row access (bracket notation):
	       FRAPPE_GET("Sales Invoice", "SINV-0001", "items[1].amount")
	       → returns the `amount` field of the 1st row in the items child table
	       The row index is 1-based. Returns None if row does not exist.

	Returns:
	    {"value": <field_value>}
	"""
	import re

	frappe.has_permission(doctype, "read", throw=True)

	# ── Pattern 3: child table bracket notation — "items[2].rate" ────────────
	child_match = re.match(r"^(\w+)\[(\d+)\]\.(\w+)$", fieldname)
	if child_match:
		child_field = child_match.group(1)
		row_idx = int(child_match.group(2)) - 1  # 1-based → 0-based
		child_fieldname = child_match.group(3)

		meta = frappe.get_meta(doctype)
		df = meta.get_field(child_field)
		if not df or df.fieldtype not in ("Table", "Table MultiSelect"):
			frappe.throw(_(f"'{child_field}' is not a Table field on {doctype}"))

		doc = frappe.get_doc(doctype, name)
		rows = doc.get(child_field) or []
		if row_idx < 0 or row_idx >= len(rows):
			return {"value": None}
		return {"value": rows[row_idx].get(child_fieldname)}

	# ── Pattern 2: Link-field traversal — "employee.ctc" ─────────────────────
	if "." in fieldname:
		parts = fieldname.split(".")
		current_doctype = doctype
		current_name = name

		for part in parts[:-1]:  # all but last are link fields to traverse
			_validate_fieldname(current_doctype, part)
			meta = frappe.get_meta(current_doctype)
			df = meta.get_field(part)
			if not df or df.fieldtype != "Link":
				frappe.throw(_(f"'{part}' is not a Link field on {current_doctype}"))
			current_name = frappe.db.get_value(current_doctype, current_name, part)
			if not current_name:
				return {"value": None}
			current_doctype = df.options
			frappe.has_permission(current_doctype, "read", throw=True)

		final_field = parts[-1]
		_validate_fieldname(current_doctype, final_field)
		return {"value": frappe.db.get_value(current_doctype, current_name, final_field)}

	# ── Pattern 1: plain field ────────────────────────────────────────────────
	_validate_fieldname(doctype, fieldname)
	return {"value": frappe.db.get_value(doctype, name, fieldname)}


# ── Grid save helpers ─────────────────────────────────────────────────────────
#
# Frappe's own link check (Document._validate_links) has a gap: when a Link field
# has "fetch from" fields (for example Student Applicant.program fills class_name),
# get_invalid_links() reads the target with as_dict=True, gets None for a missing
# record, and skips the "invalid link" entry. A non-existent value is then saved.
# The grid needs one check that works for every Link field of every DocType.

# Field types that carry no value of their own and must never be written by the grid.
_NON_VALUE_FIELDTYPES = frozenset(
	{"Section Break", "Column Break", "Tab Break", "HTML", "Button", "Heading", "Fold"}
)
# Table fields are edited through the child-table panel, not through bulk_set_value.
_TABLE_FIELDTYPES = frozenset({"Table", "Table MultiSelect"})
# Standard columns the grid may not change through bulk_set_value.
_PROTECTED_FIELDS = frozenset(
	{"name", "owner", "creation", "modified", "modified_by", "docstatus", "idx", "doctype"}
)


def _check_writable_fields(meta, fields: dict) -> None:
	"""Refuse a field the grid is not allowed to write.

	The grid already blocks these in the browser. This is the same rule on the
	server, so a direct call to bulk_set_value cannot bypass it.
	"""
	for fieldname in fields:
		if fieldname in _PROTECTED_FIELDS:
			frappe.throw(_("{0} cannot be changed here.").format(frappe.bold(fieldname)))
		df = meta.get_field(fieldname)
		if not df:
			frappe.throw(_("{0} is not a field of {1}.").format(frappe.bold(fieldname), meta.name))
		if df.fieldtype in _NON_VALUE_FIELDTYPES or df.fieldtype in _TABLE_FIELDTYPES:
			frappe.throw(_("{0} cannot be changed here.").format(frappe.bold(df.label or fieldname)))
		# Read-only fields are never writable here, including fields filled by "fetch from":
		# Frappe refreshes those itself on save, and the grid shows them locked.
		if df.read_only or df.fieldtype == "Read Only":
			frappe.throw(
				_("{0} is read-only and cannot be changed here.").format(frappe.bold(df.label or fieldname))
			)


def _check_link_values(doc, fieldnames=None) -> None:
	"""Refuse a Link value that does not exist, for any Link field of any DocType.

	Frappe's own check misses Link fields that have "fetch from" fields (see above),
	so the values are checked directly. Only fields in `fieldnames` are checked when
	it is given, so old wrong values on other fields do not block an unrelated edit.
	"""
	for df in doc.meta.get_link_fields():
		if fieldnames is not None and df.fieldname not in fieldnames:
			continue
		value = doc.get(df.fieldname)
		target = df.options
		if not value or not target:
			continue
		target_meta = frappe.get_meta(target)
		if target_meta.get("is_virtual") or getattr(target_meta, "issingle", 0):
			continue
		if not frappe.db.exists(target, value):
			frappe.throw(
				_("{0}: {1} does not exist in {2}").format(
					_(df.label or df.fieldname), frappe.bold(value), _(target)
				),
				frappe.LinkValidationError,
			)


@frappe.whitelist()
def bulk_set_value(doctype: str, updates: str) -> dict:
	"""
	Update multiple existing records via the full Frappe ORM (doc.save()).

	Every write goes through validate, before_save, after_save, on_change and
	all controller methods — identical to saving from the Frappe form UI.
	doc.save() calls notify_update() internally, so realtime list_update events
	are queued automatically and fired on the final commit.

	Submitted documents are rejected with a clear error; use the amend workflow.

	Before the save, each row is checked for (a) fields the grid may not write
	(read-only, system, layout and table fields) and (b) Link values that do not
	exist. A row that fails is reported by name and the other rows still save.

	updates: JSON array of {name: str, fields: {fieldname: value, ...}}
	Returns {"errors": [{"name": ..., "error": ..., "fields": [...]}, ...]}
	— empty list = all OK. "fields" lists the fieldnames sent for that row, so the
	grid can mark the failed cells.
	"""
	frappe.has_permission(doctype, "write", throw=True)

	updates_data: list[dict] = _fast_parse(updates) if isinstance(updates, str) else (updates or [])
	meta = frappe.get_meta(doctype)
	errors: list[dict] = []
	any_saved = False

	for item in updates_data:
		name = item.get("name")
		fields = item.get("fields") or {}
		if not name or not fields:
			continue
		# One savepoint per row: if a hook fails half way, only this row is undone.
		savepoint = "ev_bulk_row"
		frappe.db.savepoint(savepoint)
		try:
			_check_writable_fields(meta, fields)
			doc = frappe.get_doc(doctype, name)
			for fieldname, value in fields.items():
				doc.set(fieldname, value)
			_check_link_values(doc, set(fields))
			# doc.save() enforces all Frappe rules natively:
			# - submitted docs: only allow_on_submit fields may be changed
			# - cancelled docs: no edits allowed
			# - validate, before_save, after_save hooks all run
			doc.save()
			any_saved = True
		except Exception as exc:
			frappe.db.rollback(save_point=savepoint)
			errors.append({"name": name, "error": str(exc), "fields": list(fields)})
			# frappe.throw() logs the message to frappe.message_log before
			# raising, so it would appear in _server_messages and trigger
			# Frappe's default msgprint alongside our EVModal. Clear it so
			# the error is shown only through our custom UI.
			frappe.message_log.clear()

	if any_saved:
		# Explicit commit flushes all after_commit=True realtime events
		# queued by doc.save() → notify_update() in the loop above.
		frappe.db.commit()

	return {"errors": errors}


@frappe.whitelist()
def bulk_create_records(doctype: str, rows: str) -> dict:
	"""
	Create multiple new Frappe documents from a JSON array of field-value dicts.

	Args:
	    doctype: Target DocType, e.g. "Customer"
	    rows:    JSON string — [{fieldname: value, ...}, ...]
	             Empty-string values are skipped (treated as unset).

	Returns:
	    {
	        "created": [{"idx": 0, "name": "CUST-001"}, ...],
	        "errors":  [{"idx": 1, "message": "Customer Name is required"}, ...]
	    }

	Security: caller must have Create permission on the DocType.
	Each document is saved with the caller's permissions (no ignore_permissions).
	"""
	frappe.has_permission(doctype, "create", throw=True)

	rows_data: list[dict] = _fast_parse(rows) if isinstance(rows, str) else (rows or [])
	created: list[dict] = []
	errors: list[dict] = []

	for idx, row in enumerate(rows_data):
		savepoint = "ev_bulk_create_row"
		frappe.db.savepoint(savepoint)
		try:
			doc = frappe.new_doc(doctype)
			for field, value in row.items():
				# Skip empty values — let Frappe apply its own defaults
				if value is not None and value != "":
					doc.set(field, value)
			_check_link_values(doc)
			doc.insert()
			created.append({"idx": idx, "name": doc.name})
		except Exception as exc:
			frappe.db.rollback(save_point=savepoint)
			errors.append({"idx": idx, "message": str(exc)})

	if created:
		frappe.db.commit()

	return {"created": created, "errors": errors}


@frappe.whitelist()
def batch_check_links(doctype: str, rows_json: str, link_fields_json: str) -> dict:
	"""
	Check which linked values in the rows don't exist in their target DocTypes.

	Args:
	    doctype:          Source DocType being imported into (for permission check).
	    rows_json:        JSON array of {fieldname: value} dicts (already mapped).
	    link_fields_json: JSON array of {fieldname, options} describing Link fields.

	Returns:
	    {"missing": {fieldname: [val1, val2, ...]}}
	    Only fieldnames with at least one missing value are included.
	"""
	frappe.has_permission(doctype, "create", throw=True)

	rows: list[dict] = _fast_parse(rows_json) if isinstance(rows_json, str) else (rows_json or [])
	link_fields: list[dict] = (
		_fast_parse(link_fields_json) if isinstance(link_fields_json, str) else (link_fields_json or [])
	)

	missing: dict[str, list[str]] = {}

	for lf in link_fields:
		fn = lf.get("fieldname")
		options_dt = lf.get("options")
		if not fn or not options_dt:
			continue

		# Collect unique non-empty values from the rows
		values = list({str(r.get(fn, "")).strip() for r in rows if r.get(fn)})
		if not values:
			continue

		try:
			existing = {
				r["name"]
				for r in frappe.get_list(
					options_dt,
					filters=[["name", "in", values]],
					fields=["name"],
					limit=len(values) + 1,
					ignore_permissions=False,
				)
			}
		except Exception:
			continue

		missing_vals = [v for v in values if v not in existing]
		if missing_vals:
			missing[fn] = missing_vals

	return {"missing": missing}


@frappe.whitelist()
def create_link_record(doctype: str, value: str) -> dict:
	"""
	Create a minimal record in a Link DocType with the given name/value.
	Determines the correct primary field from meta.autoname.

	Returns {"name": created_doc_name, "doctype": doctype}
	"""
	frappe.has_permission(doctype, "create", throw=True)

	meta = frappe.get_meta(doctype)
	doc = frappe.new_doc(doctype)
	autoname = meta.autoname or ""

	if autoname.lower().startswith("field:"):
		# e.g. "field:item_code" → set that field
		primary_field = autoname.split(":", 1)[1].strip()
		doc.set(primary_field, value)
	elif autoname.lower() in ("prompt", "name"):
		doc.name = value
	else:
		# Fallback: set first required Data field that is empty
		for f in meta.fields:
			if f.reqd and f.fieldtype == "Data" and not doc.get(f.fieldname):
				doc.set(f.fieldname, value)
				break

	# Fill remaining required fields so insert() doesn't throw MandatoryError.
	# - Data/Text fields   → use `value` as a sensible default
	# - Link fields        → use first existing record in the linked doctype
	# - Select fields      → use first option in the options list
	for f in meta.fields:
		if not f.reqd or doc.get(f.fieldname):
			continue
		if f.fieldtype in ("Data", "Small Text", "Text", "Long Text"):
			doc.set(f.fieldname, value)
		elif f.fieldtype == "Link" and f.options:
			first_val = frappe.db.get_value(f.options, {}, "name")
			if first_val:
				doc.set(f.fieldname, first_val)
		elif f.fieldtype == "Select" and f.options:
			first_opt = (f.options or "").strip().split("\n")[0]
			if first_opt:
				doc.set(f.fieldname, first_opt)

	doc.insert(ignore_permissions=False)
	frappe.db.commit()
	return {"name": doc.name, "doctype": doctype}


@frappe.whitelist()
def frappe_child_get(
	parent_doctype: str,
	parent_name: str,
	child_field: str,
	row_index: int,
	fieldname: str,
) -> dict:
	"""
	Fetch one field from a specific row of a child table.

	Args:
	    parent_doctype: e.g. "Sales Invoice"
	    parent_name:    e.g. "SINV-0001"
	    child_field:    Table fieldname on the parent, e.g. "items"
	    row_index:      1-based row number (1 = first row)
	    fieldname:      field to read from that child row, e.g. "amount"

	Returns:
	    {"value": <field_value>}   — None if row does not exist

	Example:
	    FRAPPE_CHILD_GET("Sales Invoice", "SINV-0001", "items", 1, "amount")
	    → grand total of the first line item
	"""
	frappe.has_permission(parent_doctype, "read", throw=True)

	meta = frappe.get_meta(parent_doctype)
	df = meta.get_field(child_field)
	if not df or df.fieldtype not in ("Table", "Table MultiSelect"):
		frappe.throw(_(f"'{child_field}' is not a Table field on {parent_doctype}"))

	row_idx = int(row_index) - 1  # 1-based → 0-based
	if row_idx < 0:
		frappe.throw(_("row_index must be ≥ 1"))

	doc = frappe.get_doc(parent_doctype, parent_name)
	rows = doc.get(child_field) or []
	if row_idx >= len(rows):
		return {"value": None}

	return {"value": rows[row_idx].get(fieldname)}


@frappe.whitelist()
def frappe_aggregate(
	doctype: str,
	fieldname: str | None = None,
	aggr_type: str = "sum",
	fk1=None,
	fv1=None,
	fk2=None,
	fv2=None,
	fk3=None,
	fv3=None,
	fk4=None,
	fv4=None,
) -> dict:
	"""
	Compute SUM, COUNT, or AVG over a DocType filtered by SUMIF-style pairs.

	Accepts up to four field/value filter pairs.  Operator-style keys are
	supported: ``"transaction_date >="`` → ["transaction_date", ">=", value].

	Returns:
	    {"value": <number>}
	"""
	frappe.has_permission(doctype, "read", throw=True)

	aggr_type = (aggr_type or "sum").lower()
	if aggr_type not in ("sum", "count", "avg", "max", "min"):
		frappe.throw(_("aggr_type must be 'sum', 'count', 'avg', 'max', or 'min'."))

	if aggr_type != "count" and not fieldname:
		frappe.throw(_("fieldname is required for sum/avg/max/min aggregates."))

	if aggr_type != "count" and fieldname:
		_validate_fieldname(doctype, fieldname)

	filters = _parse_filter_pairs(fk1, fv1, fk2, fv2, fk3, fv3, fk4, fv4)

	if aggr_type == "count":
		return {"value": frappe.db.count(doctype, filters)}

	rows = frappe.get_list(
		doctype,
		filters=filters,
		fields=[fieldname],
		limit=_AGGREGATE_ROW_CAP,
		ignore_permissions=False,
	)

	raw_vals = [r[fieldname] for r in rows if r.get(fieldname) is not None]

	if not raw_vals:
		return {"value": 0 if aggr_type in ("sum", "avg") else ""}

	# max/min work on strings (ISO dates sort lexicographically) — no float cast
	if aggr_type == "max":
		return {"value": max(raw_vals)}
	if aggr_type == "min":
		return {"value": min(raw_vals)}

	vals = [float(v) for v in raw_vals]

	if aggr_type == "sum":
		return {"value": sum(vals)}

	return {"value": sum(vals) / len(vals)}


_AGG_BATCH_TTL = 30  # seconds — formula cells feel live; Redis not hammered on every keystroke


def _agg_batch_cache_key(raw: str) -> str:
	"""Stable Redis key: user + MD5 of the raw JSON queries string."""
	import hashlib

	digest = hashlib.md5(raw.encode(), usedforsecurity=False).hexdigest()
	return f"ev_agg_batch:{frappe.session.user}:{digest}"


def _invalidate_agg_cache_for_doctype(doc, method=None) -> None:
	"""
	Purge all cached batch aggregate results.

	Called as a Frappe document hook (after_insert / on_update / on_cancel /
	on_trash) whenever any document is saved so formula cells referencing the
	changed doctype are recomputed on the next grid load rather than returning
	stale Redis-cached values.

	Frappe's Redis cache key space is small (only cached values, not session/
	queue data) so the prefix scan is fast in practice.
	"""
	frappe.cache().delete_keys("ev_agg_batch:*")


def _exec_pregroup(q: dict) -> list:
	"""
	Execute a pre-grouped aggregate query (client sent fv1_list format) as a
	single SQL IN-query.  Returns a list of values parallel to q["fv1_list"].

	This is called when the client has already done the grouping step and sent
	{aggr_type, doctype, fieldname, fk1, fv1_list:[...], fk2-fk4 static filters}.
	The server runs one GROUP BY query and maps results back to the input order.
	"""
	aggr_type = (q.get("aggr_type") or "sum").lower()
	doctype = str(q.get("doctype") or "")
	fieldname = str(q.get("fieldname") or "name")
	fk1 = str(q.get("fk1") or "")
	fv1_list = [str(v) for v in (q.get("fv1_list") or [])]

	if not doctype or not fk1 or not fv1_list:
		return []

	try:
		frappe.has_permission(doctype, "read", throw=True)
	except frappe.PermissionError:
		return ["#PERM_DENIED"] * len(fv1_list)

	try:
		if aggr_type not in ("count", "get"):
			_validate_fieldname(doctype, fieldname)
		if aggr_type != "get":
			_validate_fieldname(doctype, fk1)
		else:
			_validate_fieldname(doctype, fieldname)
	except Exception:
		return ["#ARG!"] * len(fv1_list)

	# ── Static tail-filters (fk2-fk4) ────────────────────────────────────────
	extra_sql = ""
	extra_params: list = []
	for fk, fv in (
		(q.get("fk2") or "", q.get("fv2")),
		(q.get("fk3") or "", q.get("fv3")),
		(q.get("fk4") or "", q.get("fv4")),
	):
		if not fk or fv is None or fv == "":
			continue
		parts = fk.rsplit(" ", 1)
		if len(parts) == 2 and parts[1].lower() in _FILTER_OPS:
			field, op = parts[0].strip(), parts[1].strip()
		else:
			field, op = fk.strip(), "="
		try:
			_validate_fieldname(doctype, field)
		except Exception:
			continue
		extra_sql += f" AND `{field}` {op} %s"
		extra_params.append(fv)

	placeholders = ", ".join(["%s"] * len(fv1_list))

	# ── "get" path ────────────────────────────────────────────────────────────
	if aggr_type == "get":
		sql = f"SELECT `name`, `{fieldname}`" f" FROM `tab{doctype}`" f" WHERE `name` IN ({placeholders})"
		try:
			rows = frappe.db.sql(sql, fv1_list, as_dict=False)
			agg_map = {str(r[0]): r[1] for r in rows}
		except Exception:
			return [""] * len(fv1_list)
		return [agg_map.get(fv1, "") for fv1 in fv1_list]

	# ── Aggregate path ────────────────────────────────────────────────────────
	if aggr_type == "sum":
		agg_expr, default = f"SUM(`{fieldname}`)", 0
	elif aggr_type == "avg":
		agg_expr, default = f"AVG(`{fieldname}`)", 0
	elif aggr_type == "max":
		agg_expr, default = f"MAX(`{fieldname}`)", ""
	elif aggr_type == "min":
		agg_expr, default = f"MIN(`{fieldname}`)", ""
	else:  # count
		agg_expr, default = "COUNT(*)", 0

	sql = (
		f"SELECT `{fk1}`, {agg_expr}"
		f" FROM `tab{doctype}`"
		f" WHERE `{fk1}` IN ({placeholders}){extra_sql}"
		f" GROUP BY `{fk1}`"
	)
	try:
		rows = frappe.db.sql(sql, fv1_list + extra_params, as_dict=False)
		agg_map = {str(r[0]): r[1] for r in rows}
	except Exception:
		return [default] * len(fv1_list)

	return [agg_map.get(fv1, default) for fv1 in fv1_list]


@frappe.whitelist()
def frappe_aggregate_batch(queries: str) -> dict:
	"""
	Batch endpoint for FRAPPE_SUM / COUNT / AVG / MAX / MIN / GET formula functions.

	Receives a JSON array of aggregate query descriptors and returns results in
	the same order.  Queries that share the same (aggr_type, doctype, fieldname,
	fk1 key, static tail-filters fk2-fk4) are merged into a single SQL
	``WHERE fk1 IN (...) GROUP BY fk1`` — converting N DB round-trips into 1.

	Results are cached in Frappe's Redis layer for _AGG_BATCH_TTL seconds.
	Cache is automatically invalidated when any document is saved (see hooks).

	Returns:
	    {"results": [value, ...]}   — parallel to the input *queries* array.
	"""
	raw: str = queries if isinstance(queries, str) else frappe.as_json(queries)
	qs: list = (
		(_fast_parse(raw) if isinstance(raw, str) else (raw or []))
		if isinstance(queries, str)
		else (queries or [])
	)
	if not qs:
		return {"results": []}

	# ── Redis cache check ────────────────────────────────────────────────────
	import time as _time

	t0 = _time.monotonic()
	cache_key = _agg_batch_cache_key(raw)
	cached = frappe.cache().get_value(cache_key)
	if cached is not None:
		cached["_cache"] = "hit"
		return cached

	results: list = [None] * len(qs)

	# ── Group queries for IN-query optimisation ─────────────────────────────
	# Group key: (aggr_type, doctype, fieldname, fk1 [simple field — no operator],
	#             fk2+fv2, fk3+fv3, fk4+fv4 static filters)
	# Only fv1 varies within a group — each row has a different FK value.
	groups: dict = {}  # group_key → [(index, fv1), ...]
	solo: list = []  # [(index, query_dict), ...] — cannot be batched

	for i, q in enumerate(qs):
		# ── Pre-grouped compact format (client sent fv1_list) ─────────────────
		# Client has already grouped by signature; results[i] is an array
		# parallel to fv1_list (not a scalar) so the client can reconstruct
		# each cell's cache entry without re-grouping on the server.
		if isinstance(q.get("fv1_list"), list):
			results[i] = _exec_pregroup(q)
			continue

		aggr_type = (q.get("aggr_type") or "sum").lower()
		doctype = str(q.get("doctype") or "")
		fieldname = str(q.get("fieldname") or "name")
		fk1 = str(q.get("fk1") or "")
		fv1 = q.get("fv1")
		fk2 = str(q.get("fk2") or "")
		fv2 = q.get("fv2")
		fk3 = str(q.get("fk3") or "")
		fv3 = q.get("fv3")
		fk4 = str(q.get("fk4") or "")
		fv4 = q.get("fv4")

		# Batch only when fk1 is a plain field name (no trailing operator like ">="
		# or "!=") and fv1 is a non-empty scalar.
		can_group = bool(fk1 and fv1 is not None and fv1 != "" and " " not in fk1.strip())

		if can_group:
			gk = (
				aggr_type,
				doctype,
				fieldname,
				fk1,
				fk2,
				str(fv2) if fv2 is not None else "",
				fk3,
				str(fv3) if fv3 is not None else "",
				fk4,
				str(fv4) if fv4 is not None else "",
			)
			groups.setdefault(gk, []).append((i, fv1))
		else:
			solo.append((i, q))

	# ── Execute groups (IN-query) ────────────────────────────────────────────
	for gk, items in groups.items():
		aggr_type, doctype, fieldname, fk1, fk2, fv2s, fk3, fv3s, fk4, fv4s = gk

		# Permission check (once per group)
		try:
			frappe.has_permission(doctype, "read", throw=True)
		except frappe.PermissionError:
			for idx, _ in items:
				results[idx] = "#PERM_DENIED"
			continue

		# Validate field names before embedding in SQL.
		# frappe.throw() raises frappe.ValidationError (a subclass of Exception),
		# NOT DoesNotExistError — catch the base Exception to be safe.
		try:
			if aggr_type not in ("count", "get"):
				_validate_fieldname(doctype, fieldname)
			if aggr_type != "get":
				_validate_fieldname(doctype, fk1)
			else:
				# "get" uses fk1="name" which is always a valid system field
				_validate_fieldname(doctype, fieldname)
		except Exception:
			for idx, _ in items:
				results[idx] = "#ARG!"
			continue

		# Map fv1 → [indices] (handles duplicate FK values across rows). Built before the "get"
		# branch below, which needs both of these.
		fv1_to_idxs: dict = {}
		for idx, fv1 in items:
			fv1_to_idxs.setdefault(str(fv1), []).append(idx)
		unique_vals = list(fv1_to_idxs.keys())

		# ── "get" path: SELECT name, fieldname FROM tab WHERE name IN (...) ──
		if aggr_type == "get":
			try:
				placeholders = ", ".join(["%s"] * len(unique_vals))
				sql = (
					f"SELECT `name`, `{fieldname}`"
					f" FROM `tab{doctype}`"
					f" WHERE `name` IN ({placeholders})"
				)
				rows = frappe.db.sql(sql, unique_vals, as_dict=False)
				agg_map = {str(r[0]): r[1] for r in rows}
			except Exception:
				for idx, fv1 in items:
					solo.append(
						(
							idx,
							{
								"aggr_type": "get",
								"doctype": doctype,
								"fieldname": fieldname,
								"fk1": "name",
								"fv1": fv1,
							},
						)
					)
				continue
			for fv1_str, idxs in fv1_to_idxs.items():
				val = agg_map.get(fv1_str, "")
				for idx in idxs:
					results[idx] = val
			continue

		if len(unique_vals) == 1:
			# Single distinct value — delegate to the existing scalar endpoint
			res = frappe_aggregate(
				doctype=doctype,
				fieldname=fieldname,
				aggr_type=aggr_type,
				fk1=fk1,
				fv1=unique_vals[0],
				fk2=fk2 or None,
				fv2=fv2s or None,
				fk3=fk3 or None,
				fv3=fv3s or None,
				fk4=fk4 or None,
				fv4=fv4s or None,
			)
			val = res.get("value", 0)
			for idx in fv1_to_idxs[unique_vals[0]]:
				results[idx] = val
			continue

		# Build IN-query SQL (fk1 and fieldname already validated — safe as identifiers)
		table = f"`tab{doctype}`"
		fk1_col = f"`{fk1}`"

		if aggr_type == "sum":
			agg_expr = f"SUM(`{fieldname}`)"
			default = 0
		elif aggr_type == "avg":
			agg_expr = f"AVG(`{fieldname}`)"
			default = 0
		elif aggr_type == "max":
			agg_expr = f"MAX(`{fieldname}`)"
			default = ""
		elif aggr_type == "min":
			agg_expr = f"MIN(`{fieldname}`)"
			default = ""
		else:  # count
			agg_expr = "COUNT(*)"
			default = 0

		# Static tail-filters (fk2-fk4).
		# Supports operator-style keys like "transaction_date >=" — split off the
		# operator before validating the fieldname and before embedding in SQL.
		extra_sql = ""
		extra_params: list = []
		for fk, fv in ((fk2, fv2s), (fk3, fv3s), (fk4, fv4s)):
			if not fk or not fv:
				continue
			parts = fk.rsplit(" ", 1)
			if len(parts) == 2 and parts[1].lower() in _FILTER_OPS:
				field, op = parts[0].strip(), parts[1].strip()
			else:
				field, op = fk.strip(), "="
			try:
				_validate_fieldname(doctype, field)
			except Exception:
				continue
			extra_sql += f" AND `{field}` {op} %s"
			extra_params.append(fv)

		placeholders = ", ".join(["%s"] * len(unique_vals))
		sql = (
			f"SELECT {fk1_col}, {agg_expr}"
			f" FROM {table}"
			f" WHERE {fk1_col} IN ({placeholders}){extra_sql}"
			f" GROUP BY {fk1_col}"
		)
		params = unique_vals + extra_params

		try:
			rows = frappe.db.sql(sql, params, as_dict=False)
			agg_map = {str(r[0]): r[1] for r in rows}
		except Exception:
			# SQL failure — fall back to individual scalar calls for this group
			for idx, q in [
				(
					idx,
					{
						"aggr_type": aggr_type,
						"doctype": doctype,
						"fieldname": fieldname,
						"fk1": fk1,
						"fv1": fv1,
						"fk2": fk2 or None,
						"fv2": fv2s or None,
						"fk3": fk3 or None,
						"fv3": fv3s or None,
						"fk4": fk4 or None,
						"fv4": fv4s or None,
					},
				)
				for idx, fv1 in items
			]:
				solo.append((idx, q))
			continue

		for fv1_str, idxs in fv1_to_idxs.items():
			val = agg_map.get(fv1_str, default)
			for idx in idxs:
				results[idx] = val

	# ── Execute solo queries (ungroupable — operator filters, empty fv1, etc.) ─
	# (also receives "get" fallbacks pushed here by the group section)
	for idx, q in solo:
		try:
			aggr_type_s = (q.get("aggr_type") or "sum").lower()
			if aggr_type_s == "get":
				# Single FRAPPE_GET fallback — use frappe.get_value
				val = frappe.get_value(
					q.get("doctype", ""),
					q.get("fv1"),  # the document name
					q.get("fieldname") or "name",
				)
				results[idx] = val if val is not None else ""
			else:
				res = frappe_aggregate(
					doctype=q.get("doctype", ""),
					fieldname=q.get("fieldname") or "name",
					aggr_type=aggr_type_s,
					fk1=q.get("fk1"),
					fv1=q.get("fv1"),
					fk2=q.get("fk2"),
					fv2=q.get("fv2"),
					fk3=q.get("fk3"),
					fv3=q.get("fv3"),
					fk4=q.get("fk4"),
					fv4=q.get("fv4"),
				)
				results[idx] = res.get("value", 0)
		except Exception:
			results[idx] = "#ERR!"

	# ── Cache and return ─────────────────────────────────────────────────────
	elapsed_ms = round((_time.monotonic() - t0) * 1000)
	out = {"results": results, "_cache": "miss", "_ms": elapsed_ms}
	frappe.cache().set_value(cache_key, out, expires_in_sec=_AGG_BATCH_TTL)
	return out


# ── Dashboard formula-column card aggregate ────────────────────────────────────


@frappe.whitelist()
def compute_card_aggregate(
	doctype: str,
	formula: str,
	col_fieldnames: str,
	filter_op: str,
	filter_val: str,
	aggregate: str,
	agg_fieldname: str | None = None,
	list_filters: str | None = None,
	extra_filters: str | None = None,
	period_start: str | None = None,
	period_end: str | None = None,
) -> dict:
	"""
	Server-side number card computation for formula columns (⚡).

	Fetches ALL rows for the DocType (no page-limit cap), evaluates the
	HyperFormula template formula for every row, applies the card filter and
	any extra plain-field filters, then returns the requested aggregate.

	Parameters
	----------
	doctype         : Source DocType (the sheet's doctype)
	formula         : HyperFormula formula template, e.g.
	                  ``=IF(FRAPPE_COUNT("Sales Order","customer",A1)>0,"Active","Inactive")``
	col_fieldnames  : JSON array of fieldnames in column order (A→[0], B→[1], …)
	filter_op       : Operator on the formula result: ``=``, ``!=``, ``contains``,
	                  ``>``, ``<``, ``>=``, ``<=``
	filter_val      : Value to compare the formula result against
	aggregate       : ``count`` | ``sum`` | ``avg`` | ``min`` | ``max``
	agg_fieldname   : DB fieldname for sum/avg/min/max (ignored for count)
	list_filters    : JSON array of Frappe filter tuples (current sheet filters)
	extra_filters   : JSON array of ``{col, op, val}`` dicts for non-formula
	                  card filters applied after formula evaluation
	"""
	import json

	_check_doctype_permission(doctype)

	col_fields: list = (
		_fast_parse(col_fieldnames) if isinstance(col_fieldnames, str) else (col_fieldnames or [])
	)
	lst_filters: list = _fast_parse(list_filters) if isinstance(list_filters, str) and list_filters else []
	extra: list = _fast_parse(extra_filters) if isinstance(extra_filters, str) and extra_filters else []

	# agg_fieldname may be a formula/virtual col key (starts with _) — never a real DB field
	agg_field = (agg_fieldname or "").strip() or None
	if agg_field and agg_field.startswith("_"):
		agg_field = None

	# Substitute PERIOD_START() / PERIOD_END() placeholders with actual dates
	# (the client resolves these via formula_manager; pass them down so the server
	# can build correct SQL filters for formulas like FRAPPE_COUNT(...,PERIOD_START(),...))
	if period_start or period_end:
		formula = formula.replace("PERIOD_START()", f'"{period_start or ""}"')
		formula = formula.replace("PERIOD_END()", f'"{period_end   or ""}"')

	# ── Determine which DB fields we need ────────────────────────────────
	import re

	fetch_fields: set = {"name"}
	for ref in re.findall(r"([A-Z]+)1\b", formula, re.IGNORECASE):
		idx = _col_letter_to_index(ref.upper())
		if idx < len(col_fields):
			fetch_fields.add(col_fields[idx])
	# agg_field is None if it was a virtual/formula col — skip it
	if agg_field:
		fetch_fields.add(agg_field)
	for ef in extra:
		col = (ef.get("col") or "").strip()
		if col and not col.startswith("_"):
			fetch_fields.add(col)

	# ── Strategy 1: SQL-first (no row fetching) ───────────────────────────
	# Parses the formula into a subquery + COUNT — 2 DB queries, zero memory.
	sql_result = _try_sql_formula_compute(
		doctype,
		formula,
		col_fields,
		filter_op,
		str(filter_val),
		aggregate,
		agg_field,
		lst_filters,
		extra,
	)
	if sql_result is not None:
		return {"value": sql_result}

	# ── Strategy 2: Row-fetch fallback (unparseable formulas) ─────────────
	# Hard cap: 50 000 rows.  Warns if data was truncated.
	_ROW_CAP = 50_000
	rows = frappe.get_all(
		doctype,
		fields=list(fetch_fields),
		filters=lst_filters,
		limit=_ROW_CAP + 1,
	)
	truncated = len(rows) > _ROW_CAP
	if truncated:
		rows = rows[:_ROW_CAP]

	# ── Evaluate HyperFormula template for every row → row["_val"] ───────
	_eval_formula_bulk(formula, rows, col_fields)

	# ── Apply formula column filter ───────────────────────────────────────
	rows = [r for r in rows if _card_match(str(r.get("_val", "")), filter_op, str(filter_val))]

	# ── Apply any extra plain-field card filters ──────────────────────────
	for ef in extra:
		col, op, val = (ef.get("col") or ""), (ef.get("op") or "="), str(ef.get("val") or "")
		if not col or col.startswith("_"):
			continue
		rows = [r for r in rows if _card_match(str(r.get(col, "") or ""), op, val)]

	# ── Compute aggregate ─────────────────────────────────────────────────
	if aggregate == "count":
		result = len(rows)
		return {"value": result, "truncated": truncated}

	if not agg_field:
		return {"value": len(rows), "truncated": truncated}

	nums = []
	for r in rows:
		try:
			nums.append(float(r.get(agg_field) or 0))
		except (TypeError, ValueError):
			pass

	if not nums:
		return {"value": 0}

	if aggregate == "sum":
		return {"value": round(sum(nums), 2), "truncated": truncated}
	if aggregate == "avg":
		return {"value": round(sum(nums) / len(nums), 2), "truncated": truncated}
	if aggregate == "min":
		return {"value": round(min(nums), 2), "truncated": truncated}
	if aggregate == "max":
		return {"value": round(max(nums), 2), "truncated": truncated}
	return {"value": len(rows), "truncated": truncated}


# ── SQL-first formula compute ──────────────────────────────────────────────────


def _try_sql_formula_compute(
	doctype,
	formula,
	col_fields,
	filter_op,
	filter_val,
	aggregate,
	agg_field,
	lst_filters,
	extra_filters,
):
	"""
	Push formula-column card computation entirely into SQL — zero row fetching.

	Supports (any combination):
	  =IF(FRAPPE_COUNT("DT", [fk,fv …], A1, …)  op  N,   "t", "f")
	  =IF(FRAPPE_SUM  ("DT","sum_field",[fk,fv …],A1,…) op N, "t","f")
	  =IF(FRAPPE_AVG  (…)  op N, "t", "f")
	  =IF(FRAPPE_GET  ("DT", A1, "field") op "val", "t", "f")
	  =IF(A1 op "val", "t", "f")          ← direct column comparison
	  =IF(A1 op "val", "t", "f")  with filter_op "!=" / "contains" / …

	Also handles filter_op other than "=" by inverting the want_true logic.

	Returns the computed value (int/float) or None (caller falls back to row-fetch).
	"""
	import re

	expr = formula.strip().lstrip("=").strip()

	# ── Require IF(condition, true_val, false_val) ────────────────────────
	parsed = _parse_if_expr(expr)
	if not parsed:
		return None
	cond_str, true_val, false_val = parsed

	# Determine which branch the card filter selects
	filter_val_s = str(filter_val).strip()
	true_val_s = str(true_val).strip()
	false_val_s = str(false_val).strip()

	if filter_op == "=":
		if filter_val_s == true_val_s:
			want_true = True
		elif filter_val_s == false_val_s:
			want_true = False
		else:
			return None
	elif filter_op == "!=":
		if filter_val_s == true_val_s:
			want_true = False
		elif filter_val_s == false_val_s:
			want_true = True
		else:
			return None
	else:
		# contains / numeric ops on formula strings — can't push to SQL
		return None

	# Collect plain (non-formula) extra card filters
	plain_extra = [
		[ef["col"], ef.get("op", "="), ef.get("val", "")]
		for ef in extra_filters
		if (ef.get("col") or "").strip() and not ef["col"].startswith("_")
	]

	# ── Dispatch by condition type ────────────────────────────────────────
	return _sql_dispatch_condition(
		cond_str,
		want_true,
		doctype,
		col_fields,
		aggregate,
		agg_field,
		lst_filters,
		plain_extra,
	)


def _sql_dispatch_condition(
	cond_str, want_true, doctype, col_fields, aggregate, agg_field, lst_filters, plain_extra
):
	"""Route a parsed IF-condition to the right SQL handler."""
	import re

	# ── FRAPPE_COUNT / FRAPPE_SUM / FRAPPE_AVG (...) op N ─────────────────
	m_agg = re.match(
		r"^(FRAPPE_COUNT|FRAPPE_SUM|FRAPPE_AVG)\s*\((.+)\)\s*([><=!]+|<>)\s*(.+)$",
		cond_str,
		re.I | re.S,
	)
	if m_agg:
		func = m_agg.group(1).upper()
		args = _parse_formula_args(m_agg.group(2))
		comp_op = m_agg.group(3).replace("<>", "!=")
		try:
			threshold = float(m_agg.group(4).strip().strip('"'))
		except (ValueError, TypeError):
			return None
		return _sql_frappe_numeric_func(
			func,
			args,
			comp_op,
			threshold,
			want_true,
			doctype,
			col_fields,
			aggregate,
			agg_field,
			lst_filters,
			plain_extra,
		)

	# ── FRAPPE_GET("DT", A1, "field") op "val" ────────────────────────────
	m_get = re.match(
		r"^FRAPPE_GET\s*\((.+)\)\s*([><=!]+|<>)\s*(.+)$",
		cond_str,
		re.I | re.S,
	)
	if m_get:
		args = _parse_formula_args(m_get.group(1))
		comp_op = m_get.group(2).replace("<>", "!=")
		val = _strip_formula_quotes(m_get.group(3).strip())
		return _sql_frappe_get(
			args,
			comp_op,
			val,
			want_true,
			doctype,
			col_fields,
			aggregate,
			agg_field,
			lst_filters,
			plain_extra,
		)

	# ── A1 op "val" — direct column comparison ────────────────────────────
	m_col = re.match(r"^([A-Z]+)1\s*([><=!]+|<>)\s*(.+)$", cond_str, re.I)
	if m_col:
		idx = _col_letter_to_index(m_col.group(1).upper())
		row_field = col_fields[idx] if idx < len(col_fields) else "name"
		comp_op = m_col.group(2).replace("<>", "!=")
		val = _strip_formula_quotes(m_col.group(3).strip())
		neg_op = {"=": "!=", "!=": "=", ">": "<=", "<": ">=", ">=": "<", "<=": ">"}
		op = comp_op if want_true else neg_op.get(comp_op, "!=")
		try:
			return _count_or_agg(
				doctype,
				lst_filters + plain_extra + [[row_field, op, val]],
				aggregate,
				agg_field,
			)
		except Exception:
			return None

	return None


def _sql_frappe_numeric_func(
	func,
	args,
	comp_op,
	threshold,
	want_true,
	doctype,
	col_fields,
	aggregate,
	agg_field,
	lst_filters,
	plain_extra,
):
	"""
	Handle IF(FRAPPE_COUNT/SUM/AVG(…) op N, t, f).
	Builds a GROUP BY query on the link DocType, resolves passing keys,
	then counts/aggregates on the base DocType with IN / NOT IN.
	"""
	if not args:
		return None

	link_doctype = _strip_formula_quotes(args[0])
	if not link_doctype:
		return None

	sum_field, arg_start = None, 1
	if func in ("FRAPPE_SUM", "FRAPPE_AVG"):
		if len(args) < 2:
			return None
		sum_field = _strip_formula_quotes(args[1])
		arg_start = 2

	# Parse fk/fv pairs — exactly one must be a dynamic cell reference (A1)
	static_link_filters: list = []
	dynamic_pair = None

	i = arg_start
	while i + 1 < len(args):
		fk = _strip_formula_quotes(args[i])
		fv_raw = args[i + 1].strip()

		col_ref = _try_col_ref(fv_raw)
		if col_ref is not None:
			if dynamic_pair is not None:
				return None  # multiple dynamic cols — can't vectorise
			rf = col_fields[col_ref] if col_ref < len(col_fields) else "name"
			dynamic_pair = (fk, rf)
		else:
			# Operator may be embedded in fk: "transaction_date >="
			import re

			fk_m = re.match(r"^(.+?)\s+(>=|<=|!=|<>|>|<|=|like)\s*$", fk, re.I)
			if fk_m:
				static_link_filters.append(
					[fk_m.group(1).strip(), fk_m.group(2).strip(), _strip_formula_quotes(fv_raw)]
				)
			else:
				static_link_filters.append([fk, "=", _strip_formula_quotes(fv_raw)])
		i += 2

	if dynamic_pair is None:
		return None

	link_field, row_field = dynamic_pair

	# ── Query 1: GROUP BY → per-key numeric value ─────────────────────────
	try:
		if func == "FRAPPE_COUNT":
			rows = frappe.db.get_all(
				link_doctype,
				fields=[link_field, "count(*) as _v"],
				filters=static_link_filters,
				group_by=link_field,
			)
			val_map = {str(r.get(link_field, "")): (r.get("_v") or 0) for r in rows}

		elif func == "FRAPPE_SUM" and sum_field:
			rows = frappe.db.get_all(
				link_doctype,
				fields=[link_field, f"sum({sum_field}) as _v"],
				filters=static_link_filters,
				group_by=link_field,
			)
			val_map = {str(r.get(link_field, "")): float(r.get("_v") or 0) for r in rows}

		elif func == "FRAPPE_AVG" and sum_field:
			rows = frappe.db.get_all(
				link_doctype,
				fields=[link_field, f"sum({sum_field}) as _s", "count(*) as _c"],
				filters=static_link_filters,
				group_by=link_field,
			)
			val_map = {
				str(r.get(link_field, "")): float(r.get("_s") or 0) / max(int(r.get("_c") or 1), 1)
				for r in rows
			}
		else:
			return None
	except Exception:
		return None

	# ── Classify keys: passing / failing / absent ─────────────────────────
	passing_keys = {k for k, v in val_map.items() if _num_compare(v, comp_op, threshold)}
	failing_keys = {k for k in val_map if k not in passing_keys}
	zero_passes = _num_compare(0, comp_op, threshold)

	return _sql_apply_in_not_in(
		doctype,
		row_field,
		passing_keys,
		failing_keys,
		zero_passes,
		want_true,
		lst_filters,
		plain_extra,
		aggregate,
		agg_field,
	)


def _sql_frappe_get(
	args, comp_op, val, want_true, doctype, col_fields, aggregate, agg_field, lst_filters, plain_extra
):
	"""
	Handle IF(FRAPPE_GET("link_dt", A1, "field") op "val", t, f).
	Fetches matching names from link_dt, then IN/NOT IN on base doctype.
	"""
	if len(args) < 3:
		return None

	link_doctype = _strip_formula_quotes(args[0])
	fv_raw = args[1].strip()
	target_field = _strip_formula_quotes(args[2])

	col_ref = _try_col_ref(fv_raw)
	if col_ref is None:
		return None
	row_field = col_fields[col_ref] if col_ref < len(col_fields) else "name"

	# Query 1: find link_dt records where target_field op val
	try:
		matching = frappe.get_all(
			link_doctype,
			fields=["name"],
			filters=[[target_field, comp_op, val]],
			limit=0,
		)
		matching_names = {str(r["name"]) for r in matching}
	except Exception:
		return None

	passing_keys = matching_names
	failing_keys: set = set()  # we only know names of matching; non-matching = absence
	# For FRAPPE_GET: absent rows have no value → treat as "no match" (False branch),
	# so zero_passes is passed as False below (absent = no match).
	return _sql_apply_in_not_in(
		doctype,
		row_field,
		passing_keys,
		failing_keys,
		False,
		want_true,
		lst_filters,
		plain_extra,
		aggregate,
		agg_field,
	)


def _sql_apply_in_not_in(
	doctype,
	row_field,
	passing_keys,
	failing_keys,
	zero_passes,
	want_true,
	lst_filters,
	plain_extra,
	aggregate,
	agg_field,
):
	"""
	Build final base-DocType filters from passing/failing key sets and run aggregate.

	zero_passes = True  means rows ABSENT from the link table also satisfy the condition.
	want_true   = True  means we want rows where condition is True.
	"""
	base = lst_filters + plain_extra

	try:
		if want_true:
			if zero_passes:
				# True = passing_keys union (all rows not in val_map)
				# Equivalent: exclude failing_keys
				if failing_keys:
					return _count_or_agg(
						doctype, [*base, [row_field, "not in", list(failing_keys)]], aggregate, agg_field
					)
				else:
					return _count_or_agg(doctype, base, aggregate, agg_field)
			else:
				# True = only passing_keys
				if not passing_keys:
					return 0
				return _count_or_agg(
					doctype, [*base, [row_field, "in", list(passing_keys)]], aggregate, agg_field
				)
		else:
			if zero_passes:
				# False = rows in val_map that don't pass = failing_keys
				if not failing_keys:
					return 0
				return _count_or_agg(
					doctype, [*base, [row_field, "in", list(failing_keys)]], aggregate, agg_field
				)
			else:
				# False = failing_keys union (rows absent from val_map)
				# = total - passing_keys
				total = frappe.db.count(doctype, filters=base)
				if not passing_keys:
					return total
				passing_count = frappe.db.count(
					doctype, filters=[*base, [row_field, "in", list(passing_keys)]]
				)
				return total - passing_count
	except Exception:
		return None


def _count_or_agg(doctype, filters, aggregate, agg_field):
	"""Run COUNT or numeric aggregate on doctype with given filters."""
	if aggregate == "count" or not agg_field:
		return frappe.db.count(doctype, filters=filters)
	matched = frappe.get_all(doctype, fields=[agg_field], filters=filters, limit=0)
	nums = [float(r.get(agg_field) or 0) for r in matched if r.get(agg_field) is not None]
	if not nums:
		return 0
	if aggregate == "sum":
		return round(sum(nums), 2)
	if aggregate == "avg":
		return round(sum(nums) / len(nums), 2)
	if aggregate == "min":
		return round(min(nums), 2)
	if aggregate == "max":
		return round(max(nums), 2)
	return len(nums)


# ── compute_card_aggregate helpers ────────────────────────────────────────────


def _col_letter_to_index(letters: str) -> int:
	"""'A' → 0, 'B' → 1, 'Z' → 25, 'AA' → 26"""
	result = 0
	for c in letters.upper():
		result = result * 26 + (ord(c) - ord("A") + 1)
	return result - 1


def _card_match(cell: str, op: str, val: str) -> bool:
	cell, val = cell.strip(), val.strip()
	if op == "=":
		return cell == val
	if op == "!=":
		return cell != val
	if op == "contains":
		return val.lower() in cell.lower()
	try:
		c, v = float(cell), float(val)
		if op == ">":
			return c > v
		if op == "<":
			return c < v
		if op == ">=":
			return c >= v
		if op == "<=":
			return c <= v
	except (ValueError, TypeError):
		pass
	return False


def _eval_formula_bulk(formula: str, rows: list, col_fields: list) -> None:
	"""
	Evaluate a HyperFormula formula template for every row.
	Sets row["_val"] in-place.

	Supports nested IFs of arbitrary depth, FRAPPE_FUNC comparisons on
	both sides of a condition, and bare FRAPPE_FUNC / column references.
	"""
	import re

	expr = formula.strip()
	if expr.startswith("="):
		expr = expr[1:].strip()

	# ── IF(condition, true_val, false_val) — handles nested IFs recursively
	parsed_if = _parse_if_expr(expr)
	if parsed_if:
		cond_str, true_val_str, false_val_str = parsed_if
		conds = _eval_condition_bulk(cond_str, rows, col_fields)
		true_rows = [row for row, c in zip(rows, conds, strict=False) if c]
		false_rows = [row for row, c in zip(rows, conds, strict=False) if not c]
		if true_rows:
			_eval_branch_value(true_val_str, true_rows, col_fields)
		if false_rows:
			_eval_branch_value(false_val_str, false_rows, col_fields)
		return

	# ── Bare FRAPPE_* returning a number ─────────────────────────────────
	m = re.match(r"^(FRAPPE_COUNT|FRAPPE_SUM|FRAPPE_AVG|FRAPPE_GET)\s*\((.+)\)$", expr, re.I | re.S)
	if m:
		func, args_str = m.group(1).upper(), m.group(2)
		nums = _frappe_func_bulk(func, _parse_formula_args(args_str), rows, col_fields)
		for row, v in zip(rows, nums, strict=False):
			row["_val"] = v
		return

	# ── Bare column reference A1 ─────────────────────────────────────────
	col_m = re.match(r"^([A-Z]+)1$", expr, re.I)
	if col_m:
		idx = _col_letter_to_index(col_m.group(1).upper())
		field = col_fields[idx] if idx < len(col_fields) else "name"
		for row in rows:
			row["_val"] = str(row.get(field, "") or "")
		return

	# ── Fallback ─────────────────────────────────────────────────────────
	for row in rows:
		row["_val"] = ""


def _eval_branch_value(val_str: str, rows: list, col_fields: list) -> None:
	"""
	Evaluate one branch of an IF expression (may be a literal, nested IF,
	or FRAPPE_FUNC call) and set row["_val"] for each row in the subset.
	"""
	import re

	val_str = val_str.strip()

	# Nested IF — recurse
	if re.match(r"^IF\s*\(", val_str, re.I):
		_eval_formula_bulk("=" + val_str, rows, col_fields)
		return

	# FRAPPE_FUNC returning a value
	fm = re.match(r"^(FRAPPE_COUNT|FRAPPE_SUM|FRAPPE_AVG|FRAPPE_GET)\s*\((.+)\)$", val_str, re.I | re.S)
	if fm:
		func = fm.group(1).upper()
		args = _parse_formula_args(fm.group(2))
		nums = _frappe_func_bulk(func, args, rows, col_fields)
		for row, v in zip(rows, nums, strict=False):
			row["_val"] = v
		return

	# Literal string / number
	literal = _strip_formula_quotes(val_str)
	for row in rows:
		row["_val"] = literal


def _parse_if_expr(expr: str):
	"""
	Parse ``IF(cond, true, false)`` → (cond_str, true_val, false_val) or None.
	"""
	import re

	if not re.match(r"^IF\s*\(", expr, re.I):
		return None
	# Strip leading "IF("
	inner = re.sub(r"^IF\s*\(\s*", "", expr, flags=re.I)
	# Remove trailing ")"
	if inner.endswith(")"):
		inner = inner[:-1]
	parts = _parse_formula_args(inner)
	if len(parts) < 3:
		return None
	return (
		parts[0].strip(),
		_strip_formula_quotes(parts[1].strip()),
		_strip_formula_quotes(parts[2].strip()),
	)


def _eval_condition_bulk(cond_str: str, rows: list, col_fields: list) -> list:
	"""Evaluate a condition string for all rows, returning list[bool]."""
	import re

	# FRAPPE_FUNC(...) op threshold
	fm = re.match(
		r"^(FRAPPE_COUNT|FRAPPE_SUM|FRAPPE_AVG|FRAPPE_GET)\s*\((.+)\)\s*([><=!]+|<>)\s*(.+)$",
		cond_str,
		re.I | re.S,
	)
	if fm:
		func = fm.group(1).upper()
		args = _parse_formula_args(fm.group(2))
		op = fm.group(3).replace("<>", "!=")
		rhs = fm.group(4).strip()
		lhs_nums = _frappe_func_bulk(func, args, rows, col_fields)

		# RHS may be another FRAPPE_FUNC (e.g. FRAPPE_SUM(...) >= FRAPPE_SUM(...))
		rhs_fm = re.match(
			r"^(FRAPPE_COUNT|FRAPPE_SUM|FRAPPE_AVG|FRAPPE_GET)\s*\((.+)\)$",
			rhs,
			re.I | re.S,
		)
		if rhs_fm:
			rhs_func = rhs_fm.group(1).upper()
			rhs_args = _parse_formula_args(rhs_fm.group(2))
			rhs_nums = _frappe_func_bulk(rhs_func, rhs_args, rows, col_fields)
			return [_num_compare(l, op, r) for l, r in zip(lhs_nums, rhs_nums, strict=False)]

		try:
			threshold = float(rhs.strip('"'))
		except (ValueError, TypeError):
			threshold = 0.0
		return [_num_compare(v, op, threshold) for v in lhs_nums]

	# Column reference op literal  (A1 = "value")
	cm = re.match(r"^([A-Z]+)1\s*([><=!]+|<>)\s*(.+)$", cond_str, re.I)
	if cm:
		idx = _col_letter_to_index(cm.group(1).upper())
		field = col_fields[idx] if idx < len(col_fields) else "name"
		op = cm.group(2).replace("<>", "!=")
		val = _strip_formula_quotes(cm.group(3).strip())
		return [_card_match(str(r.get(field, "") or ""), op, val) for r in rows]

	return [True] * len(rows)


def _frappe_func_bulk(func: str, args: list, rows: list, col_fields: list) -> list:
	"""
	Evaluate FRAPPE_COUNT / FRAPPE_SUM / FRAPPE_AVG for all rows at once
	using a single GROUP BY query when possible.
	"""
	import re

	if not args:
		return [0] * len(rows)

	link_doctype = _strip_formula_quotes(args[0])
	if not link_doctype:
		return [0] * len(rows)

	# FRAPPE_SUM / FRAPPE_AVG: 2nd arg is the sum field
	sum_field = None
	arg_start = 1
	if func in ("FRAPPE_SUM", "FRAPPE_AVG"):
		if len(args) < 2:
			return [0] * len(rows)
		sum_field = _strip_formula_quotes(args[1])
		arg_start = 2

	# Parse fk/fv pairs — fv may be a column ref (A1) or a string literal
	static_filters: list = []
	dynamic_pairs: list = []  # (link_field, row_fieldname)

	i = arg_start
	while i + 1 < len(args):
		fk = _strip_formula_quotes(args[i])
		fv_raw = args[i + 1].strip()

		# Handle operator-embedded fieldnames: "transaction_date >=" → field="transaction_date", op=">="
		fk_op = "="
		fk_m = re.match(r"^(.+?)\s+(>=|<=|!=|<>|>|<|=|like)\s*$", fk, re.I)
		if fk_m:
			fk, fk_op = fk_m.group(1).strip(), fk_m.group(2).strip()

		col_ref = _try_col_ref(fv_raw)
		if col_ref is not None:
			rf = col_fields[col_ref] if col_ref < len(col_fields) else "name"
			dynamic_pairs.append((fk, rf))
		else:
			static_filters.append([fk, fk_op, _strip_formula_quotes(fv_raw)])
		i += 2

	# All-static: same value for every row
	if not dynamic_pairs:
		try:
			cnt = frappe.db.count(link_doctype, filters=static_filters) or 0
		except Exception:
			cnt = 0
		return [cnt] * len(rows)

	# Single dynamic field: one GROUP BY query covers all rows
	if len(dynamic_pairs) == 1:
		dyn_fk, dyn_rf = dynamic_pairs[0]
		row_vals = [r.get(dyn_rf) for r in rows]
		unique = list({v for v in row_vals if v is not None})
		if not unique:
			return [0] * len(rows)

		bulk_filters = [*static_filters, [dyn_fk, "in", unique]]
		val_map: dict = {}
		try:
			if func == "FRAPPE_COUNT":
				agg_rows = frappe.db.get_all(
					link_doctype,
					fields=[dyn_fk, "count(*) as cnt"],
					filters=bulk_filters,
					group_by=dyn_fk,
				)
				val_map = {str(r.get(dyn_fk, "")): (r.get("cnt") or 0) for r in agg_rows}
			elif func == "FRAPPE_SUM" and sum_field:
				agg_rows = frappe.db.get_all(
					link_doctype,
					fields=[dyn_fk, f"sum({sum_field}) as s"],
					filters=bulk_filters,
					group_by=dyn_fk,
				)
				val_map = {str(r.get(dyn_fk, "")): float(r.get("s") or 0) for r in agg_rows}
			elif func == "FRAPPE_AVG" and sum_field:
				agg_rows = frappe.db.get_all(
					link_doctype,
					fields=[dyn_fk, f"sum({sum_field}) as s", "count(*) as cnt"],
					filters=bulk_filters,
					group_by=dyn_fk,
				)
				val_map = {
					str(r.get(dyn_fk, "")): (float(r.get("s") or 0) / max(int(r.get("cnt") or 1), 1))
					for r in agg_rows
				}
		except Exception:
			pass

		return [val_map.get(str(rv), 0) for rv in row_vals]

	# Multiple dynamic fields: row-by-row (rare)
	results = []
	for row in rows:
		dyn_f = [[fk, "=", row.get(rf)] for fk, rf in dynamic_pairs]
		all_f = static_filters + dyn_f
		try:
			if func == "FRAPPE_COUNT":
				v: float = float(frappe.db.count(link_doctype, filters=all_f) or 0)
			elif func in ("FRAPPE_SUM", "FRAPPE_AVG") and sum_field:
				r_rows = frappe.db.get_all(link_doctype, fields=[sum_field], filters=all_f)
				nums = [float(x[sum_field] or 0) for x in r_rows if x.get(sum_field) is not None]
				v = sum(nums) if func == "FRAPPE_SUM" else (sum(nums) / len(nums) if nums else 0.0)
			else:
				v = 0.0
		except Exception:
			v = 0.0
		results.append(v)
	return results


def _parse_formula_args(args_str: str) -> list:
	"""Split formula args on commas, respecting quoted strings and nested parens."""
	parts, current, depth, in_q = [], [], 0, False
	for ch in args_str:
		if ch == '"' and depth == 0:
			in_q = not in_q
			current.append(ch)
		elif not in_q:
			if ch == "(":
				depth += 1
				current.append(ch)
			elif ch == ")":
				depth -= 1
				current.append(ch)
			elif ch == "," and depth == 0:
				parts.append("".join(current))
				current = []
			else:
				current.append(ch)
		else:
			current.append(ch)
	if current:
		parts.append("".join(current))
	return parts


def _strip_formula_quotes(s: str) -> str:
	s = s.strip()
	if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
		return s[1:-1]
	return s


def _try_col_ref(s: str):
	"""Return column index if s matches A1, B1, AA1 etc., else None."""
	import re

	m = re.match(r"^([A-Z]+)1$", s.strip(), re.I)
	return _col_letter_to_index(m.group(1).upper()) if m else None


def _num_compare(val, op: str, threshold: float) -> bool:
	try:
		v = float(val)
		if op == ">":
			return v > threshold
		if op == "<":
			return v < threshold
		if op == ">=":
			return v >= threshold
		if op == "<=":
			return v <= threshold
		if op == "=":
			return v == threshold
		if op == "!=":
			return v != threshold
	except (TypeError, ValueError):
		pass
	return False


@frappe.whitelist()
def gl_balance(
	account: str,
	company: str,
	from_date: str | None = None,
	to_date: str | None = None,
	cost_center: str | None = None,
	finance_book: str | None = None,
) -> dict:
	"""
	Return the net GL balance (debit - credit) for an account/company/period.

	Used by the ``GL_BALANCE(account, company, …)`` HyperFormula function.
	Returns ``{"value": 0, "note": "…"}`` if ERPNext is not installed so the
	formula degrades gracefully on vanilla Frappe setups.

	Args:
	    account:      GL account name (e.g. "Cash - ACME").
	    company:      Company name.
	    from_date:    Start of period (inclusive).  None = no lower bound.
	    to_date:      End of period (inclusive).    None = no upper bound.
	    cost_center:  Optional cost-centre filter.
	    finance_book: Optional finance-book filter (also matches IS NULL rows).

	Returns:
	    {"value": <Decimal as float>}
	"""
	if not frappe.db.table_exists("tabGL Entry"):
		return {"value": 0, "note": "GL Entry not available — ERPNext not installed"}

	frappe.has_permission("GL Entry", "read", throw=True)

	conditions = [
		"account  = %(account)s",
		"company  = %(company)s",
		"is_cancelled = 0",
	]
	params: dict = {"account": account, "company": company}

	if from_date:
		conditions.append("posting_date >= %(from_date)s")
		params["from_date"] = from_date
	if to_date:
		conditions.append("posting_date <= %(to_date)s")
		params["to_date"] = to_date
	if cost_center:
		conditions.append("cost_center = %(cost_center)s")
		params["cost_center"] = cost_center
	if finance_book:
		conditions.append("(finance_book = %(finance_book)s OR finance_book IS NULL)")
		params["finance_book"] = finance_book

	where = " AND ".join(conditions)
	result = frappe.db.sql(
		f"SELECT COALESCE(SUM(debit), 0) - COALESCE(SUM(credit), 0) AS balance "
		f"FROM `tabGL Entry` WHERE {where}",
		params,
		as_dict=True,
	)

	return {"value": float(result[0].get("balance") or 0)}


@frappe.whitelist()
def stock_qty(
	item_code: str,
	warehouse: str,
	as_of_date: str | None = None,
) -> dict:
	"""
	Return current or point-in-time stock quantity for an item/warehouse.

	Used by ``STOCK_QTY(item_code, warehouse [, as_of_date])``.
	When *as_of_date* is omitted the Bin table is used (fast current balance).
	When a date is supplied the Stock Ledger Entry table is queried to return
	the running total up to that date.

	Returns:
	    {"value": <float>}
	"""
	if not frappe.db.table_exists("tabBin"):
		return {"value": 0, "note": "Bin doctype not available — ERPNext not installed"}

	frappe.has_permission("Bin", "read", throw=True)

	if not as_of_date:
		qty = frappe.db.get_value(
			"Bin",
			{"item_code": item_code, "warehouse": warehouse},
			"actual_qty",
		)
		return {"value": float(qty or 0)}

	# Point-in-time balance via Stock Ledger Entry
	result = frappe.db.sql(
		"""
        SELECT COALESCE(SUM(actual_qty), 0) AS qty
        FROM   `tabStock Ledger Entry`
        WHERE  item_code    = %(item)s
          AND  warehouse    = %(wh)s
          AND  posting_date <= %(date)s
          AND  docstatus    = 1
        """,
		{"item": item_code, "wh": warehouse, "date": as_of_date},
		as_dict=True,
	)
	return {"value": float(result[0].get("qty") or 0)}


@frappe.whitelist()
def item_price(
	item_code: str,
	price_list: str,
	qty: float | None = None,
	customer: str | None = None,
	uom: str | None = None,
) -> dict:
	"""
	Return the selling price for an item from an Item Price record.

	Used by ``ITEM_PRICE(item_code, price_list [, qty, customer, uom])``.
	Picks the most recently valid price (valid_from ≤ today ≤ valid_upto).
	Falls back to the most recently modified price if no date-valid record
	exists.

	Args:
	    item_code:  Item code.
	    price_list: Price list name (e.g. "Standard Selling").
	    qty:        Quantity (reserved for future tiered pricing — unused now).
	    customer:   Customer (reserved for future customer-group pricing).
	    uom:        Unit of measure filter.

	Returns:
	    {"value": <float>}
	"""
	if not frappe.db.table_exists("tabItem Price"):
		return {"value": 0, "note": "Item Price not available — ERPNext not installed"}

	frappe.has_permission("Item Price", "read", throw=True)

	filters: dict = {
		"item_code": item_code,
		"price_list": price_list,
		"selling": 1,
	}
	if uom:
		filters["uom"] = uom

	prices = frappe.get_list(
		"Item Price",
		filters=filters,
		fields=["price_list_rate", "valid_from", "valid_upto"],
		order_by="valid_from desc",
	)

	if not prices:
		return {"value": 0}

	today = frappe.utils.today()

	for p in prices:
		valid_from = p.get("valid_from")
		valid_upto = p.get("valid_upto")

		# Skip if period hasn't started yet
		if valid_from and frappe.utils.getdate(valid_from) > frappe.utils.getdate(today):
			continue
		# Skip if period has already ended
		if valid_upto and frappe.utils.getdate(valid_upto) < frappe.utils.getdate(today):
			continue

		return {"value": float(p.get("price_list_rate") or 0)}

	# Fallback: first record (most recent valid_from regardless of date range)
	return {"value": float(prices[0].get("price_list_rate") or 0)}


# ── Read ──────────────────────────────────────────────────────────────────────


@frappe.whitelist()
def get_workbooks(doctype_name: str) -> list[dict]:
	"""
	Return workbooks for *doctype_name* that the current user may open:
	  - workbooks they own, OR
	  - workbooks marked is_public = 1 (by any user).

	Returns a lightweight list (no heavy JSON fields) suitable for rendering
	the "Open View" dialog.
	"""
	frappe.has_permission(doctype_name, "read", throw=True)

	# Raw SQL is cleaner than chaining frappe.get_all OR-filters.
	return frappe.db.sql(
		"""
		SELECT name, title, owner, is_public, modified
		FROM   `tabExcel Workbook`
		WHERE  doctype_name = %(dt)s
		  AND  (owner = %(user)s OR is_public = 1)
		ORDER  BY modified DESC
		LIMIT  200
		""",
		{"dt": doctype_name, "user": frappe.session.user},
		as_dict=True,
	)


@frappe.whitelist()
def load_workbook(name: str) -> dict:
	"""
	Return the full workbook document (including JSON config fields).

	Returns {"not_found": True} when the workbook has been deleted so the
	client can silently clear its stale user_settings reference instead of
	showing a scary "Not found" error dialog.
	"""
	try:
		doc = frappe.get_doc("Excel Workbook", name)
	except frappe.DoesNotExistError:
		return {"not_found": True, "name": name}

	if doc.owner != frappe.session.user and not doc.is_public:
		frappe.throw(
			_("You don't have permission to open this view."),
			frappe.PermissionError,
		)

	return doc.as_dict()


# ── Write ─────────────────────────────────────────────────────────────────────


@frappe.whitelist()
def save_workbook(
	title: str,
	doctype_name: str,
	columns_config: str,
	formula_columns: str,
	filters: str,
	sort_by: str | None = None,
	is_public: int = 0,
	workbook_name: str | None = None,
	join_config: str | None = None,
	sheets: str | None = None,
	chart_overlays: str | None = None,
	format_store: str | None = None,
	cond_fmt_rules: str | None = None,
	view_state: str | None = None,
) -> dict:
	"""
	Create a new workbook or update an existing one.
	"""
	frappe.has_permission(doctype_name, "read", throw=True)

	if workbook_name:
		doc = frappe.get_doc("Excel Workbook", workbook_name)
		if doc.owner != frappe.session.user and not frappe.has_role("System Manager"):
			frappe.throw(_("You can only update your own saved views."))
	else:
		doc = frappe.new_doc("Excel Workbook")
		doc.doctype_name = doctype_name

	doc.title = title
	doc.is_public = frappe.utils.cint(is_public)
	doc.columns_config = columns_config
	doc.formula_columns = formula_columns
	doc.filters = filters
	doc.sort_by = sort_by or "{}"
	doc.join_config = join_config or "{}"
	doc.sheets = sheets or "[]"
	doc.chart_overlays = chart_overlays or "[]"
	doc.format_store = format_store or "{}"
	doc.cond_fmt_rules = cond_fmt_rules or "[]"
	doc.view_state = view_state or "{}"

	doc.save(ignore_permissions=False)

	return {"name": doc.name, "title": doc.title}


# ── Delete ────────────────────────────────────────────────────────────────────


@frappe.whitelist()
def delete_workbook(name: str) -> dict:
	"""
	Delete a workbook.  Only the owner or a System Manager may do this.
	The controller's before_delete hook also enforces this, but we check
	early here to give a clear error message.
	"""
	doc = frappe.get_doc("Excel Workbook", name)

	if doc.owner != frappe.session.user and not frappe.has_role("System Manager"):
		frappe.throw(_("You can only delete your own saved views."))

	frappe.delete_doc("Excel Workbook", name, ignore_permissions=True)

	return {"success": True}


# ── V2.4.5 — IntelliFlow AI: helpers + 4-layer validation + discovery ─────────


# ── Shared helpers ────────────────────────────────────────────────────────────


def _sample_and_profile(doctype: str, field: str, limit: int = 100) -> dict:
	"""
	Sample `limit` values from doctype.field and detect their structural pattern.

	Returns:
	  {
	    "values":     set[str]           — sampled raw values
	    "raw_list":   list[str]          — ordered list (for cardinality check)
	    "pattern":    str                — one of:
	                    "naming_series"  (EMP-0001, SINV-2024-00001)
	                    "hash"           (len≥10, no spaces)
	                    "email"          (contains @)
	                    "date"           (YYYY-MM-DD…)
	                    "numeric"        (pure numbers / decimals)
	                    "text"           (everything else)
	    "prefix_set": set[str]           — naming-series prefixes (empty otherwise)
	  }
	"""
	import re

	raw = [str(r) for r in frappe.get_all(doctype, pluck=field, limit=limit) if r]
	if not raw:
		return {"values": set(), "raw_list": [], "pattern": "text", "prefix_set": set()}

	s = raw[:50]
	n = len(s)

	def _ratio(fn):
		return sum(1 for v in s if fn(v)) / n

	series_re = re.compile(r"^([A-Z][A-Z0-9-]*)-(?:\d{4}-)?[0-9]+$")

	if _ratio(lambda v: bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", v))) >= 0.8:
		return {"values": set(raw), "raw_list": raw, "pattern": "email", "prefix_set": set()}

	if _ratio(lambda v: bool(re.match(r"^\d{4}-\d{2}-\d{2}", v))) >= 0.8:
		return {"values": set(raw), "raw_list": raw, "pattern": "date", "prefix_set": set()}

	if _ratio(lambda v: bool(re.match(r"^-?\d+\.?\d*$", v))) >= 0.8:
		return {"values": set(raw), "raw_list": raw, "pattern": "numeric", "prefix_set": set()}

	series_matches = [series_re.match(v) for v in s]
	if sum(1 for m in series_matches if m) / n >= 0.7:
		prefixes = {m.group(1) for m in series_matches if m}
		return {"values": set(raw), "raw_list": raw, "pattern": "naming_series", "prefix_set": prefixes}

	if _ratio(lambda v: len(v) >= 10 and " " not in v) >= 0.8:
		return {"values": set(raw), "raw_list": raw, "pattern": "hash", "prefix_set": set()}

	return {"values": set(raw), "raw_list": raw, "pattern": "text", "prefix_set": set()}


def _l1_pattern_score(src_prof: dict, tgt_prof: dict) -> float:
	"""Layer 1 score from structural pattern comparison."""
	sp = src_prof["pattern"]
	tp = tgt_prof["pattern"]

	if sp == "naming_series" and tp == "naming_series":
		# Boost if they share at least one naming-series prefix
		if src_prof["prefix_set"] & tgt_prof["prefix_set"]:
			return 0.95
		return 0.80

	if sp == "hash" and tp == "hash":
		return 0.70
	if sp == "email" and tp == "email":
		return 0.60
	if sp == tp:
		return 0.40  # same generic type
	return 0.10  # different


# Hard incompatible pattern pairs — type-gate (Layer 2)
_INCOMPATIBLE_PAIRS = {
	("date", "numeric"),
	("numeric", "date"),
	("date", "email"),
	("email", "date"),
	("numeric", "email"),
	("email", "numeric"),
	("numeric", "naming_series"),
	("naming_series", "numeric"),
	("numeric", "hash"),
	("hash", "numeric"),
}


def _grade(composite: float, method: str) -> str:
	if method == "meta":
		return "S"
	if composite >= 0.80:
		return "A"
	if composite >= 0.60:
		return "B"
	if composite >= 0.40:
		return "C"
	if composite >= 0.25:
		return "D"
	return "F"


def _cardinality_and_coverage(src_list: list, tgt_list: list):
	"""Returns (cardinality_str, coverage_float)."""
	from collections import Counter

	src_set = set(src_list)
	tgt_set = set(tgt_list)
	coverage = round(len(src_set & tgt_set) / max(len(src_set), 1), 2)
	tgt_cnt = Counter(tgt_list)
	matched = [v for v in src_list if v in tgt_set]
	max_m = max((tgt_cnt.get(v, 0) for v in matched), default=0)
	return ("1:1" if max_m <= 1 else "1:N"), coverage


# ── Link-graph cache (V2.4.5) ─────────────────────────────────────────────────
#
# Instead of calling frappe.get_meta(dt) for every DocType in a loop
# (= N individual DB round-trips), we run ONE SQL JOIN on tabDocField
# that returns every Link field across the entire site, then cache the
# result in Redis for 5 minutes.
#
# Performance:
#   Before: suggest_joins ≈ 300 get_meta() calls  ≈ 300 DB queries
#   After : 1-2 SQL JOIN queries → edge list       → cached, no repeat

_GRAPH_CACHE_KEY = "ev_link_graph_edges_v3"  # v3: includes Dynamic Link
_GRAPH_CACHE_TTL = 300  # seconds
_DYNAMIC_LINK_CACHE_KEY = "ev_dynamic_link_edges_v1"
_DYNAMIC_LINK_SAMPLE_LIMIT = 500  # Sample size for Dynamic Link detection


def _get_all_link_edges() -> list:
	"""
	Return every Link-field edge across all non-single DocTypes in a SINGLE SQL
	query. Child tables (istable=1) are now included as edge SOURCES so that
	canvas users can join child-table DocTypes (e.g. "Timesheet Detail") that
	carry a Link field to the base DocType. The JOIN target is still restricted
	to non-child, non-single DocTypes.

	Custom fields (tabCustom Field) are merged via UNION ALL.
	Result is cached in Redis for 5 min; subsequent calls are instant.

	Each entry: {
	  "doctype": str, "fieldname": str, "label": str, "target": str,
	  "is_child_src": int   # 1 when the source DocType is a child table
	}
	"""
	cached = frappe.cache().get_value(_GRAPH_CACHE_KEY)
	if cached is not None:
		return cached

	sql = """
		SELECT df.parent      AS doctype,
		       df.fieldname,
		       COALESCE(NULLIF(df.label, ''), df.fieldname) AS label,
		       df.options     AS target,
		       src.istable    AS is_child_src
		FROM   `tabDocField` df
		JOIN   `tabDocType`  src ON src.name = df.parent
		JOIN   `tabDocType`  tgt ON tgt.name = df.options
		WHERE  df.fieldtype = 'Link'
		  AND  df.options IS NOT NULL AND df.options != ''
		  AND  src.issingle = 0
		  AND  tgt.issingle = 0 AND tgt.istable = 0

		UNION ALL

		SELECT cf.dt          AS doctype,
		       cf.fieldname,
		       COALESCE(NULLIF(cf.label, ''), cf.fieldname) AS label,
		       cf.options     AS target,
		       src.istable    AS is_child_src
		FROM   `tabCustom Field` cf
		JOIN   `tabDocType`  src ON src.name = cf.dt
		JOIN   `tabDocType`  tgt ON tgt.name = cf.options
		WHERE  cf.fieldtype = 'Link'
		  AND  cf.options IS NOT NULL AND cf.options != ''
		  AND  src.issingle = 0
		  AND  tgt.issingle = 0 AND tgt.istable = 0
	"""
	edges = [dict(row) for row in frappe.db.sql(sql, as_dict=True)]

	# Add Dynamic Link edges (polymorphic relationships)
	dynamic_edges = _get_dynamic_link_edges()
	edges.extend(dynamic_edges)

	frappe.cache().set_value(_GRAPH_CACHE_KEY, edges, expires_in_sec=_GRAPH_CACHE_TTL)
	return edges


def _get_dynamic_link_edges() -> list:
	"""
	Detect Dynamic Link relationships - polymorphic joins where target is determined
	by another field's value.

	Dynamic Link detection uses the `options` field to find the reference field,
	then determines targets based on that field's type:

	Case 1 - Select field (deterministic):
	  party_type (Select: "Customer\nSupplier\nEmployee")
	  party (Dynamic Link → party_type)
	  → Targets: Customer, Supplier, Employee (parsed from Select options)

	Case 2 - Link to DocType (data sampling):
	  reference_doctype (Link → DocType)
	  reference_name (Dynamic Link → reference_doctype)
	  → Sample data to discover which DocTypes are actually used

	Case 3 - Link with set_query (data sampling):
	  link_doctype (Link → DocType, controller restricts via set_query)
	  link_name (Dynamic Link → link_doctype)
	  → Sample data to discover actual usage

	Common in: Comment, Address, Contact, File, Version, Communication, Payment Entry

	Returns edges with metadata:
	  {
	    "doctype": source DocType,
	    "fieldname": Dynamic Link field,
	    "label": field label,
	    "target": discovered target DocType,
	    "is_child_src": 0 or 1,
	    "is_dynamic": True,
	    "ref_field": reference field name,
	    "ref_field_type": "Select" or "Link",
	    "detection_method": "select" or "data_sample",
	    "sample_count": count (None for Select, int for sampled)
	  }
	"""
	cached = frappe.cache().get_value(_DYNAMIC_LINK_CACHE_KEY)
	if cached is not None:
		return cached

	# Step 1: Find all Dynamic Link fields
	dynamic_fields_sql = """
		SELECT df.parent      AS doctype,
		       df.fieldname,
		       COALESCE(NULLIF(df.label, ''), df.fieldname) AS label,
		       df.options     AS ref_field,
		       src.istable    AS is_child_src
		FROM   `tabDocField` df
		JOIN   `tabDocType`  src ON src.name = df.parent
		WHERE  df.fieldtype = 'Dynamic Link'
		  AND  df.options IS NOT NULL AND df.options != ''
		  AND  src.issingle = 0

		UNION ALL

		SELECT cf.dt          AS doctype,
		       cf.fieldname,
		       COALESCE(NULLIF(cf.label, ''), cf.fieldname) AS label,
		       cf.options     AS ref_field,
		       src.istable    AS is_child_src
		FROM   `tabCustom Field` cf
		JOIN   `tabDocType`  src ON src.name = cf.dt
		WHERE  cf.fieldtype = 'Dynamic Link'
		  AND  cf.options IS NOT NULL AND cf.options != ''
		  AND  src.issingle = 0
	"""
	dynamic_fields = frappe.db.sql(dynamic_fields_sql, as_dict=True)

	if not dynamic_fields:
		frappe.cache().set_value(_DYNAMIC_LINK_CACHE_KEY, [], expires_in_sec=_GRAPH_CACHE_TTL)
		return []

	# Step 2: For each Dynamic Link field, determine targets based on ref field type
	edges = []

	for df in dynamic_fields:
		try:
			# Find the reference field's meta to determine how to get targets
			meta = frappe.get_meta(df["doctype"])
			ref_field_obj = meta.get_field(df["ref_field"])

			if not ref_field_obj:
				continue  # Reference field doesn't exist

			targets_data = []

			# Case 1: Select field - parse options directly (no data sampling needed!)
			if ref_field_obj.fieldtype == "Select" and ref_field_obj.options:
				# Options are newline-separated: "Customer\nSupplier\nEmployee"
				select_options = [opt.strip() for opt in ref_field_obj.options.split("\n") if opt.strip()]

				# Each option is a potential target DocType
				for target_dt in select_options:
					# First check if DocType exists in tabDocType (fast DB check)
					exists = frappe.db.exists("DocType", target_dt)
					if not exists:
						continue  # DocType doesn't exist - skip silently

					# Now safely get meta (we know it exists)
					try:
						target_meta = frappe.get_meta(target_dt)
						if not target_meta.issingle and not target_meta.istable:
							targets_data.append(
								{
									"target_doctype": target_dt,
									"count": None,  # No count for Select (all are possible)
									"method": "select",
								}
							)
					except Exception:
						continue  # Meta fetch failed - skip

			# Case 2 & 3: Link field (to DocType or restricted) - sample actual data
			elif ref_field_obj.fieldtype == "Link":
				sample_sql = """
					SELECT `{ref_field}` AS target_doctype, COUNT(*) AS count
					FROM `tab{doctype}`
					WHERE `{ref_field}` IS NOT NULL AND `{ref_field}` != ''
					GROUP BY `{ref_field}`
					LIMIT {limit}
				""".format(doctype=df["doctype"], ref_field=df["ref_field"], limit=_DYNAMIC_LINK_SAMPLE_LIMIT)

				sampled_targets = frappe.db.sql(sample_sql, as_dict=True)

				for row in sampled_targets:
					# First check if DocType exists (fast DB check)
					exists = frappe.db.exists("DocType", row["target_doctype"])
					if not exists:
						continue  # DocType doesn't exist - skip silently

					# Now safely get meta
					try:
						target_meta = frappe.get_meta(row["target_doctype"])
						if not target_meta.issingle and not target_meta.istable:
							targets_data.append(
								{
									"target_doctype": row["target_doctype"],
									"count": row["count"],
									"method": "data_sample",
								}
							)
					except Exception:
						continue

				# Create edges for discovered targets
				for target_info in targets_data:
					edges.append(
						{
							"doctype": df["doctype"],
							"fieldname": df["fieldname"],
							"label": df["label"],
							"target": target_info["target_doctype"],
							"is_child_src": df["is_child_src"],
							"is_dynamic": True,
							"ref_field": df["ref_field"],
							"ref_field_type": ref_field_obj.fieldtype,
							"detection_method": target_info["method"],
							"sample_count": target_info.get("count"),
						}
					)

		except Exception:
			# If table doesn't exist or meta fetch fails, skip this Dynamic Link
			continue

	frappe.cache().set_value(_DYNAMIC_LINK_CACHE_KEY, edges, expires_in_sec=_GRAPH_CACHE_TTL)
	return edges


# ── Main validation endpoint ──────────────────────────────────────────────────


def _csv_to_headers_rows(text: str):
	"""Parse CSV text → {"headers": [...], "rows": [[...]]}."""
	import csv
	import io

	reader = csv.reader(io.StringIO(text))
	all_rows = list(reader)
	if not all_rows:
		return {"headers": [], "rows": []}
	return {"headers": all_rows[0], "rows": all_rows[1:]}


def _validate_external_url(url: str) -> None:
	"""Block SSRF: reject requests to RFC-1918 / loopback / link-local addresses."""
	import ipaddress
	import re
	import socket
	from urllib.parse import urlparse

	parsed = urlparse(url)
	if parsed.scheme not in ("http", "https"):
		frappe.throw(_("Only http/https URLs are allowed"))

	hostname = parsed.hostname or ""

	# Block raw IP addresses that are private / loopback / link-local
	try:
		ip = ipaddress.ip_address(hostname)
		if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved:
			frappe.throw(_("Requests to internal network addresses are not allowed"))
	except ValueError:
		# It's a hostname — resolve and check
		block_patterns = re.compile(
			r"^(localhost|.*\.local|.*\.internal|.*\.intranet|metadata\.google\.internal)$",
			re.I,
		)
		if block_patterns.match(hostname):
			frappe.throw(_("Requests to internal hostnames are not allowed"))
		try:
			resolved_ip = ipaddress.ip_address(socket.gethostbyname(hostname))
			if resolved_ip.is_private or resolved_ip.is_loopback or resolved_ip.is_link_local:
				frappe.throw(_("Requests to internal network addresses are not allowed"))
		except OSError:
			frappe.throw(_(f"Cannot resolve hostname: {hostname}"))


@frappe.whitelist()
def fetch_google_sheet(url: str, tab_name: str = ""):
	"""Fetch a public Google Sheet by URL and return headers + rows.

	Uses the CSV export endpoint — no API key needed for sheets shared
	as "Anyone with link can view".
	"""
	import re

	import requests

	match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
	if not match:
		return {"error": "Invalid Google Sheets URL — could not find spreadsheet ID."}

	sheet_id = match.group(1)
	export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
	if tab_name:
		from urllib.parse import quote

		export_url += f"&sheet={quote(tab_name)}"

	try:
		resp = requests.get(export_url, timeout=20, allow_redirects=True)
		resp.raise_for_status()
		ct = resp.headers.get("content-type", "")
		if "text/html" in ct:
			return {"error": "Could not access sheet. Ensure it is shared as 'Anyone with link can view'."}
		return _csv_to_headers_rows(resp.text)
	except Exception as e:
		return {"error": str(e)}


@frappe.whitelist()
def fetch_url_text(url: str):
	"""Return raw text content of any URL (used by CSV / JSON import)."""
	import requests

	_validate_external_url(url)
	try:
		resp = requests.get(url, timeout=20)
		resp.raise_for_status()
		return {"text": resp.text}
	except Exception as e:
		return {"error": str(e)}


@frappe.whitelist()
def fetch_web_api(url: str, method: str = "GET", headers: str = "", json_path: str = ""):
	"""Call a REST endpoint and normalise the JSON response to headers + rows.

	Args:
		url:        Endpoint URL.
		method:     HTTP method (GET / POST).
		headers:    JSON-encoded dict of request headers.
		json_path:  Dot-notation path into the JSON response (e.g. "data.items").
	"""
	import json

	import requests

	_validate_external_url(url)
	hdrs: dict = {}
	if headers:
		try:
			hdrs = _fast_parse(headers) if isinstance(headers, str) else dict(headers)
		except Exception:
			pass

	try:
		fn = getattr(requests, method.lower(), requests.get)
		resp = fn(url, headers=hdrs, timeout=25)
		resp.raise_for_status()
		data = resp.json()
	except Exception as e:
		return {"error": str(e)}

	# Traverse dot-notation path
	if json_path:
		for key in json_path.split("."):
			if isinstance(data, dict):
				data = data.get(key)
			elif isinstance(data, list) and key.isdigit():
				data = data[int(key)]
			else:
				data = None
			if data is None:
				return {"error": f"Path '{json_path}' not found in the response."}

	if not isinstance(data, list):
		data = [data] if isinstance(data, dict) else []
	if not data:
		return {"headers": [], "rows": []}

	first = data[0]
	col_keys: list = list(first.keys()) if isinstance(first, dict) else [f"col{i}" for i in range(len(first))]
	rows = []
	for item in data:
		if isinstance(item, dict):
			rows.append([item.get(k, "") for k in col_keys])
		elif isinstance(item, list):
			rows.append(item)
		else:
			rows.append([item])

	return {"headers": col_keys, "rows": rows}


@frappe.whitelist()
def extract_pdf_tables(pdf_b64: str):
	"""Extract all tables from a base64-encoded PDF using pdfplumber.

	Returns {"tables": [{"headers": [...], "rows": [[...], ...]}, ...]}.
	"""
	import base64
	import io

	try:
		import pdfplumber
	except ImportError:
		return {"error": "pdfplumber is not installed. Run: pip install pdfplumber"}

	try:
		pdf_bytes = base64.b64decode(pdf_b64)
		tables = []
		with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
			for page in pdf.pages:
				for tbl in page.extract_tables() or []:
					if not tbl:
						continue
					headers = [str(c or f"Col{i + 1}") for i, c in enumerate(tbl[0])]
					rows = [[str(cell or "") for cell in row] for row in tbl[1:]]
					if rows:
						tables.append({"headers": headers, "rows": rows})
		return {"tables": tables}
	except Exception as e:
		return {"error": str(e)}


# ── Field-type compatibility helpers ─────────────────────────────────────────

_NUMERIC_FT: frozenset = frozenset({"Currency", "Float", "Int", "Percent"})
_TEXT_FT: frozenset = frozenset({"Data", "Link", "Select", "Small Text", "Text", "Long Text"})
_DATE_FT: frozenset = frozenset({"Date", "Datetime", "Time"})


def _field_type_compat(ft1: str, ft2: str) -> float:
	"""Confidence multiplier when fieldtypes differ.  1.0 = perfect, 0.65 = incompatible."""
	if not ft1 or not ft2:
		return 1.0
	if ft1 == ft2:
		return 1.0
	if ft1 in _NUMERIC_FT and ft2 in _NUMERIC_FT:
		return 1.0
	if ft1 in _TEXT_FT and ft2 in _TEXT_FT:
		return 0.95
	if ft1 in _DATE_FT and ft2 in _DATE_FT:
		return 0.95
	return 0.65  # incompatible types (e.g. Currency ↔ Link)


@frappe.whitelist()
def smart_lookup_suggest(
	source_doctype: str,
	target_doctype: str,
	source_headers: str,
	target_headers: str,
	source_sample: str = "[]",
	target_sample: str = "[]",
):
	"""AI-based 4-layer join column suggestion (no LLM).

	Layer 0: Primary Key Match — ≥50 % of source-col values are found as `name` IDs
	         in the target sample (foreign-key / primary-key detection).
	Layer 1: Frappe meta — Link fields on source DocType pointing to target DocType.
	Layer 2a: Exact fieldname match.
	Layer 2b: Fuzzy label / fieldname similarity (rapidfuzz token_sort_ratio).
	Layer 3: Data-content match — Polars-accelerated Jaccard similarity on sampled
	         unique values (Python-set fallback when Polars unavailable).

	Data-type guardrails: confidence is multiplied by a compatibility factor when
	source and target fieldtypes are semantically incompatible (e.g. Float ↔ Link).

	Returns: list of {source_col, target_col, strategy, confidence, reason}
	sorted by confidence descending, one card per (unclaimed) source column.
	"""
	# Permission check — user must be able to read both DocTypes
	if source_doctype:
		try:
			frappe.has_permission(source_doctype, "read", throw=True)
		except Exception:
			frappe.throw(_(f"You do not have read access to {source_doctype}"))
	if target_doctype:
		try:
			frappe.has_permission(target_doctype, "read", throw=True)
		except Exception:
			frappe.throw(_(f"You do not have read access to {target_doctype}"))

	import json

	from rapidfuzz import fuzz

	try:
		import polars as pl

		_USE_POLARS = True
	except ImportError:
		_USE_POLARS = False

	src_headers = _fast_parse(source_headers)  # [{fieldname, label, fieldtype, options}, …]
	tgt_headers = _fast_parse(target_headers)
	src_sample = _fast_parse(source_sample)  # [[val, …], …]  ≤ 200 rows
	tgt_sample = _fast_parse(target_sample)

	suggestions: list = []
	claimed_src: set = set()
	claimed_tgt: set = set()

	# ── Layer 1: Frappe meta — Link fields pointing to target DocType ──────────
	if source_doctype and target_doctype:
		try:
			src_meta = frappe.get_meta(source_doctype)
			for df in src_meta.fields:
				if df.fieldtype == "Link" and df.options == target_doctype:
					if df.fieldname not in claimed_src:
						suggestions.append(
							{
								"source_col": df.fieldname,
								"target_col": "name",
								"strategy": "link_field",
								"confidence": 0.97,
								"reason": (
									f"'{df.label or df.fieldname}' is a Link field to {target_doctype}"
								),
							}
						)
						claimed_src.add(df.fieldname)
						claimed_tgt.add("name")
		except Exception:
			frappe.clear_messages()

	# ── Layer 2a: Exact fieldname match (runs BEFORE Primary Key Match) ────────
	# Structural equality trumps data-based heuristics.
	claimed_src_l2 = {s["source_col"] for s in suggestions}
	claimed_tgt_l2 = {s["target_col"] for s in suggestions}

	for sf in src_headers:
		if sf["fieldname"] in claimed_src_l2:
			continue
		for tf in tgt_headers:
			if tf["fieldname"] in claimed_tgt_l2:
				continue
			if sf["fieldname"] == tf["fieldname"]:
				compat = _field_type_compat(sf.get("fieldtype", ""), tf.get("fieldtype", ""))
				suggestions.append(
					{
						"source_col": sf["fieldname"],
						"target_col": tf["fieldname"],
						"strategy": "header_match",
						"confidence": round(0.98 * compat, 3),
						"reason": f"Exact fieldname match: '{sf['fieldname']}'",
					}
				)
				claimed_src_l2.add(sf["fieldname"])
				claimed_tgt_l2.add(tf["fieldname"])
				break

	# Merge back before Layer 0 so it respects exact-fieldname claims
	claimed_src = {s["source_col"] for s in suggestions}
	claimed_tgt = {s["target_col"] for s in suggestions}

	# ── Layer 0c: Doctype-name FK detection (structural, no data needed) ────────
	# Catches: IGA.customer + target_doctype="Customer" → JOIN IGA.customer→Customer.name
	# AND:    Customer.name + source_doctype="Customer" → JOIN Customer.name→report.customer
	#
	# Pure structural signal: column fieldname == snake_case(doctype) is a definitive FK.
	# Works even when the report is filtered (low sample overlap) or when there's no meta.

	def _doctype_fk_bonus(si_idx, ti_idx):
		"""Return a 0-0.05 data-overlap bonus to add to structural base confidence."""
		if not src_sample or not tgt_sample:
			return 0.0
		src_v = {str(r[si_idx]) for r in src_sample if si_idx < len(r) and str(r[si_idx]).strip()}
		tgt_v = {str(r[ti_idx]) for r in tgt_sample if ti_idx < len(r) and str(r[ti_idx]).strip()}
		if not src_v or not tgt_v:
			return 0.0
		ratio = len(src_v & tgt_v) / len(src_v)
		return round(ratio * 0.05, 3)

	# 0c-A: source column named like target_doctype → FK to target.name
	if target_doctype and "name" not in claimed_tgt:
		dt_snake = frappe.scrub(target_doctype)  # "Customer" → "customer"
		name_ti = next((i for i, tf in enumerate(tgt_headers) if tf["fieldname"] == "name"), None)
		if name_ti is not None:
			for si, sf in enumerate(src_headers):
				if sf["fieldname"] in claimed_src:
					continue
				fn = sf["fieldname"]
				if fn == dt_snake or fn == dt_snake + "_id":
					bonus = _doctype_fk_bonus(si, name_ti)
					suggestions.append(
						{
							"source_col": fn,
							"target_col": "name",
							"strategy": "primary_key_match",
							"confidence": round(min(0.99, 0.94 + bonus), 3),
							"reason": (
								f"'{sf.get('label') or fn}' column name matches DocType "
								f"'{target_doctype}' — foreign key on {target_doctype}.name"
							),
						}
					)
					claimed_src.add(fn)
					claimed_tgt.add("name")
					break

	# 0c-B: source.name (primary key) → target column named like source_doctype
	if source_doctype and "name" not in claimed_src:
		dt_snake = frappe.scrub(source_doctype)
		src_name_si = next((i for i, sf in enumerate(src_headers) if sf["fieldname"] == "name"), None)
		if src_name_si is not None:
			for ti, tf in enumerate(tgt_headers):
				if tf["fieldname"] in claimed_tgt:
					continue
				fn = tf["fieldname"]
				if fn == dt_snake or fn == dt_snake + "_id":
					bonus = _doctype_fk_bonus(src_name_si, ti)
					suggestions.append(
						{
							"source_col": "name",
							"target_col": fn,
							"strategy": "primary_key_match",
							"confidence": round(min(0.99, 0.94 + bonus), 3),
							"reason": (
								f"'{tf.get('label') or fn}' column name matches source DocType "
								f"'{source_doctype}' — join on {source_doctype}.name → {fn}"
							),
						}
					)
					claimed_src.add("name")
					claimed_tgt.add(fn)
					break

	# Refresh claims after Layer 0c
	claimed_src = {s["source_col"] for s in suggestions}
	claimed_tgt = {s["target_col"] for s in suggestions}

	# ── Layer 0a: Primary Key Match — source values → target `name` column ─────
	# Skip source columns that already matched via exact fieldname (Layer 2a).
	# Threshold lowered to 0.3: a filtered report may show only a subset of rows
	# but the join is still correct (e.g. 6/20 IGA customers appear in Customer sample).
	if src_sample and tgt_sample:
		name_ti = next((i for i, tf in enumerate(tgt_headers) if tf["fieldname"] == "name"), None)
		if name_ti is not None and "name" not in claimed_tgt:
			tgt_names_raw = [
				str(r[name_ti]) for r in tgt_sample if name_ti < len(r) and str(r[name_ti]).strip()
			]
			if tgt_names_raw:
				tgt_name_set = set(tgt_names_raw)
				if _USE_POLARS:
					tgt_name_pl = pl.Series(tgt_names_raw).unique()

				for si, sf in enumerate(src_headers):
					if sf["fieldname"] in claimed_src:
						continue
					src_raw = [str(r[si]) for r in src_sample if si < len(r) and str(r[si]).strip()]
					if not src_raw:
						continue

					if _USE_POLARS:
						src_pl = pl.Series(src_raw).unique()
						overlap = int(src_pl.is_in(tgt_name_pl).sum())
						total = len(src_pl)
					else:
						src_set = set(src_raw)
						overlap = len(src_set & tgt_name_set)
						total = len(src_set)

					if total == 0:
						continue
					ratio = overlap / total
					if ratio >= 0.3:
						# Scale: 0.3 ratio→~0.82, 1.0 ratio→0.99
						confidence = round(min(0.99, 0.72 + ratio * 0.27), 3)
						suggestions.append(
							{
								"source_col": sf["fieldname"],
								"target_col": "name",
								"strategy": "primary_key_match",
								"confidence": confidence,
								"reason": (
									f"{overlap}/{total} source values match target IDs "
									f"({int(ratio * 100)}% hit-rate)"
								),
							}
						)
						claimed_src.add(sf["fieldname"])
						claimed_tgt.add("name")
						break

	# ── Layer 0b: Source `name` (ID) values → any unclaimed target column ──────
	# Covers doctype→report joins where the report stores IDs in a non-`name` col
	# (e.g. Customer.name = "CUST-001" appears in Report.customer_name column).
	if src_sample and tgt_sample and "name" not in claimed_src:
		src_name_si = next((i for i, sf in enumerate(src_headers) if sf["fieldname"] == "name"), None)
		if src_name_si is not None:
			src_raw = [
				str(r[src_name_si])
				for r in src_sample
				if src_name_si < len(r) and str(r[src_name_si]).strip()
			]
			if src_raw:
				src_set = set(src_raw)
				best_j, best_tf, best_cnt = 0.0, None, 0
				for ti, tf in enumerate(tgt_headers):
					# Layer 0b intentionally ignores claimed_tgt:
					# `name → FK` is a distinct join pattern; multiple source columns
					# pointing to the same target is fine — user picks which to apply.
					if tf["fieldname"] == "name":
						continue
					tgt_raw = [str(r[ti]) for r in tgt_sample if ti < len(r) and str(r[ti]).strip()]
					if not tgt_raw:
						continue
					tgt_set = set(tgt_raw)
					inter = len(src_set & tgt_set)
					ratio = inter / len(src_set) if src_set else 0.0
					if ratio >= 0.3 and ratio > best_j:
						best_j, best_tf, best_cnt = ratio, tf, inter

				if best_tf:
					suggestions.append(
						{
							"source_col": "name",
							"target_col": best_tf["fieldname"],
							"strategy": "primary_key_match",
							"confidence": round(min(0.96, 0.82 + best_j * 0.14), 3),
							"reason": (
								f"{best_cnt}/{len(src_set)} source IDs match "
								f"'{best_tf.get('label') or best_tf['fieldname']}' values "
								f"({int(best_j * 100)}% hit-rate)"
							),
						}
					)
					claimed_src.add("name")
					claimed_tgt.add(best_tf["fieldname"])

	# ── Layer 2b: Fuzzy label + fieldname similarity ───────────────────────────
	claimed_src_l2 = {s["source_col"] for s in suggestions}
	claimed_tgt_l2 = {s["target_col"] for s in suggestions}

	for sf in src_headers:
		if sf["fieldname"] in claimed_src_l2:
			continue
		best_score, best_tf = 0, None
		for tf in tgt_headers:
			if tf["fieldname"] in claimed_tgt_l2:
				continue
			src_label = (sf.get("label") or sf["fieldname"]).lower()
			tgt_label = (tf.get("label") or tf["fieldname"]).lower()
			lbl_score = fuzz.token_sort_ratio(src_label, tgt_label)
			fn_score = fuzz.ratio(
				sf["fieldname"].replace("_", " "),
				tf["fieldname"].replace("_", " "),
			)
			score = max(lbl_score, fn_score)
			if score > best_score:
				best_score, best_tf = score, tf
		if best_tf and best_score >= 75:
			compat = _field_type_compat(sf.get("fieldtype", ""), best_tf.get("fieldtype", ""))
			suggestions.append(
				{
					"source_col": sf["fieldname"],
					"target_col": best_tf["fieldname"],
					"strategy": "header_match",
					"confidence": round((best_score / 100) * compat, 3),
					"reason": (
						f"'{sf.get('label') or sf['fieldname']}' ≈ "
						f"'{best_tf.get('label') or best_tf['fieldname']}' "
						f"({best_score}% similarity)"
					),
				}
			)
			claimed_src_l2.add(sf["fieldname"])
			claimed_tgt_l2.add(best_tf["fieldname"])

	# Merge back
	claimed_src = {s["source_col"] for s in suggestions}
	claimed_tgt = {s["target_col"] for s in suggestions}

	# ── Layer 3: Data-content match (Polars Jaccard) ───────────────────────────
	if src_sample and tgt_sample:
		for si, sf in enumerate(src_headers):
			if sf["fieldname"] in claimed_src:
				continue
			src_raw = [str(r[si]) for r in src_sample if si < len(r) and str(r[si]).strip()]
			if not src_raw:
				continue

			best_j, best_tf, best_cnt = 0.0, None, 0
			for ti, tf in enumerate(tgt_headers):
				if tf["fieldname"] in claimed_tgt:
					continue
				tgt_raw = [str(r[ti]) for r in tgt_sample if ti < len(r) and str(r[ti]).strip()]
				if not tgt_raw:
					continue

				if _USE_POLARS:
					src_set = set(pl.Series(src_raw).unique().to_list())
					tgt_set = set(pl.Series(tgt_raw).unique().to_list())
				else:
					src_set = set(src_raw)
					tgt_set = set(tgt_raw)

				inter = len(src_set & tgt_set)
				union = len(src_set | tgt_set)
				j = inter / union if union else 0.0
				if j > best_j:
					best_j, best_tf, best_cnt = j, tf, inter

			if best_tf and best_j >= 0.2:
				compat = _field_type_compat(sf.get("fieldtype", ""), best_tf.get("fieldtype", ""))
				suggestions.append(
					{
						"source_col": sf["fieldname"],
						"target_col": best_tf["fieldname"],
						"strategy": "data_content_match",
						"confidence": round(best_j * compat, 3),
						"reason": (
							f"{best_cnt} shared unique values " f"({int(best_j * 100)}% Jaccard overlap)"
						),
					}
				)
				claimed_src.add(sf["fieldname"])
				claimed_tgt.add(best_tf["fieldname"])

	suggestions.sort(key=lambda x: -x["confidence"])
	seen, result = set(), []
	for s in suggestions:
		if s["source_col"] not in seen:
			seen.add(s["source_col"])
			result.append(s)

	return result


# ── Schema graph & relational auto-expansion ──────────────────────────────────


@frappe.whitelist()
def expand_relationship(
	source_doctype: str,
	target_doctype: str,
	join_field: str,
	source_values: str,
	agg_preset: str = "latest",
	return_fields: str = "[]",
) -> dict:
	"""Fetch related rows from *target_doctype* for the given *source_values*.

	Handles two cardinalities automatically:
	  N:1 — join_field is on source side; target's `name` is the match key.
	  1:N — target has a Link field pointing back to source_doctype.

	agg_preset: "latest" | "sum" | "count" | "average"
	source_values: JSON list of join-key values from the source sheet (≤ 500).
	return_fields: JSON list of fieldnames to pull (auto-selected if empty).

	Uses Polars when available for sub-500 ms processing at 100k-row scale.

	Returns: {columns, rows, cardinality, agg_preset} | {error, columns, rows}
	"""
	import json

	_check_doctype_permission(source_doctype)
	_check_doctype_permission(target_doctype)

	values = [str(v) for v in _fast_parse(source_values) if v]
	fields_wanted = _fast_parse(return_fields)
	if not values:
		return {"columns": [], "rows": [], "cardinality": "unknown"}

	try:
		import polars as pl

		_USE_POLARS = True
	except ImportError:
		_USE_POLARS = False

	tgt_meta = frappe.get_meta(target_doctype)
	col_label = {df.fieldname: (df.label or df.fieldname) for df in tgt_meta.fields}

	# Auto-detect cardinality: does target have a Link field back to source?
	reverse_field = next(
		(df.fieldname for df in tgt_meta.fields if df.fieldtype == "Link" and df.options == source_doctype),
		None,
	)
	cardinality = "1:N" if reverse_field else "N:1"

	# Auto-select return fields when not specified
	if not fields_wanted:
		fields_wanted = [
			df.fieldname
			for df in tgt_meta.fields
			if df.fieldtype in ("Currency", "Float", "Int", "Data", "Link", "Select")
			and not df.hidden
			and getattr(df, "in_list_view", 0)
		][:6]
	if not fields_wanted:
		fields_wanted = [
			df.fieldname
			for df in tgt_meta.fields
			if df.fieldtype in ("Currency", "Float", "Int") and not df.hidden
		][:5]

	placeholders = ", ".join(["%s"] * len(values))

	try:
		# ── 1:N aggregation presets ────────────────────────────────────────────
		if cardinality == "1:N" and agg_preset in ("count", "sum", "average"):
			numeric_fields = [
				df.fieldname
				for df in tgt_meta.fields
				if df.fieldtype in ("Currency", "Float", "Int")
				and not df.hidden
				and (not fields_wanted or df.fieldname in fields_wanted)
			][:5]

			if agg_preset == "count":
				sql = f"""
					SELECT `{reverse_field}` AS _src_key, COUNT(name) AS `count`
					FROM `tab{target_doctype}`
					WHERE `{reverse_field}` IN ({placeholders}) AND docstatus < 2
					GROUP BY `{reverse_field}`
				"""
				rows = frappe.db.sql(sql, values, as_dict=True)
				return {
					"columns": [
						{"fieldname": "_src_key", "label": source_doctype},
						{"fieldname": "count", "label": f"Count of {target_doctype}"},
					],
					"rows": rows,
					"cardinality": cardinality,
					"agg_preset": agg_preset,
				}

			if numeric_fields:
				pfx = "sum_" if agg_preset == "sum" else "avg_"
				fn = "SUM" if agg_preset == "sum" else "AVG"
				agg_expr = ", ".join(f"{fn}(`{f}`) AS `{pfx}{f}`" for f in numeric_fields)
				sql = f"""
					SELECT `{reverse_field}` AS _src_key, {agg_expr}
					FROM `tab{target_doctype}`
					WHERE `{reverse_field}` IN ({placeholders}) AND docstatus < 2
					GROUP BY `{reverse_field}`
				"""
				rows = frappe.db.sql(sql, values, as_dict=True)
				return {
					"columns": [
						{"fieldname": "_src_key", "label": source_doctype},
					]
					+ [
						{"fieldname": f"{pfx}{f}", "label": f"{agg_preset.title()} {col_label.get(f, f)}"}
						for f in numeric_fields
					],
					"rows": rows,
					"cardinality": cardinality,
					"agg_preset": agg_preset,
				}

		# ── N:1 or 1:N "latest" — row-level fetch ─────────────────────────────
		fld_sql = ", ".join(f"`{f}`" for f in fields_wanted) if fields_wanted else "name"
		join_col = reverse_field if cardinality == "1:N" else "name"
		sql = f"""
			SELECT `{join_col}` AS _src_key, {fld_sql}
			FROM `tab{target_doctype}`
			WHERE `{join_col}` IN ({placeholders}) AND docstatus < 2
			ORDER BY modified DESC
		"""
		rows = frappe.db.sql(sql, values, as_dict=True)

		# For 1:N "latest": keep only the most-recent row per source key
		if agg_preset == "latest" and cardinality == "1:N":
			seen_keys: set = set()
			deduped = []
			for row in rows:
				k = row.get("_src_key")
				if k not in seen_keys:
					seen_keys.add(k)
					deduped.append(row)
			rows = deduped

		# Polars round-trip — validates types and normalises nulls
		if _USE_POLARS and rows:
			rows = pl.from_dicts(rows).to_dicts()

		return {
			"columns": [{"fieldname": "_src_key", "label": join_field}]
			+ [{"fieldname": f, "label": col_label.get(f, f)} for f in fields_wanted],
			"rows": rows,
			"cardinality": cardinality,
			"agg_preset": agg_preset,
		}

	except Exception as e:
		frappe.log_error(frappe.get_traceback(), "expand_relationship error")
		return {"error": str(e), "columns": [], "rows": []}


@frappe.whitelist()
def smart_lookup_fetch(
	lookup_value: str,
	target_doctype: str,
	return_field: str,
	source_doctype: str = "",
) -> str:
	"""Scalar lookup powering the SMART_LOOKUP() HyperFormula formula.

	Strategy 1: Direct name lookup — frappe.db.get_value(target_doctype, lookup_value, field).
	Strategy 2: Reverse Link — find a Link field on target pointing to source_doctype,
	            then filter by that field.

	Returns "" when no match is found.  Always returns a string.
	"""
	if not lookup_value or not target_doctype or not return_field:
		return ""

	_check_doctype_permission(target_doctype)
	_validate_fieldname(target_doctype, return_field)

	# Strategy 1: lookup_value is the `name` (ID) in target_doctype
	try:
		val = frappe.db.get_value(target_doctype, str(lookup_value), return_field)
		if val is not None:
			return str(val)
	except Exception:
		pass

	# Strategy 2: target has a Link field back to source_doctype
	if source_doctype:
		try:
			tgt_meta = frappe.get_meta(target_doctype)
			for df in tgt_meta.fields:
				if df.fieldtype == "Link" and df.options == source_doctype:
					val = frappe.db.get_value(
						target_doctype,
						{df.fieldname: str(lookup_value)},
						return_field,
						order_by="modified desc",
					)
					if val is not None:
						return str(val)
					break
		except Exception:
			pass

	return ""


# ── V3.3 — Activity column tag helpers ────────────────────────────────────────


@frappe.whitelist()
def get_doctype_tags(doctype):
	"""Return all tags used on this DocType (distinct, sorted)."""
	frappe.has_permission(doctype, "read", throw=True)
	rows = frappe.get_all(
		"Tag Link",
		filters={"document_type": doctype},
		fields=["tag"],
		distinct=True,
		order_by="tag asc",
		limit=200,
	)
	return [r.tag for r in rows if r.tag]


@frappe.whitelist()
def add_doc_tag(doctype, docname, tag):
	"""Add a tag to a document."""
	frappe.has_permission(doctype, "write", throw=True)
	from frappe.desk.doctype.tag.tag import add_tag

	return add_tag(tag, doctype, docname)


@frappe.whitelist()
def remove_doc_tag(doctype, docname, tag):
	"""Remove a tag from a document."""
	frappe.has_permission(doctype, "write", throw=True)
	from frappe.desk.doctype.tag.tag import remove_tag

	return remove_tag(tag, doctype, docname)


def _get_active_formula_configs():
	"""Return the list of active non-system Excel Formula configs (with preset filters).
	Shared by extend_bootinfo and the realtime broadcast."""
	if not frappe.db.table_exists("Excel Formula"):
		return []
	configs = frappe.get_all(
		"Excel Formula",
		filters={"is_active": 1, "is_system": 0},
		fields=[
			"formula_name",
			"label",
			"category",
			"formula_type",
			"source_doctype",
			"target_fieldname",
			"description",
		],
		order_by="formula_name asc",
	)
	for cfg in configs:
		cfg["preset_filters"] = frappe.get_all(
			"Excel Formula Filter",
			filters={"parent": cfg["formula_name"]},
			fields=["filter_key", "filter_value"],
			order_by="idx asc",
		)
	return configs


def extend_bootinfo(bootinfo):
	"""Inject active Excel Formula configs into frappe.boot so the JS plugin
	can register dynamic HyperFormula functions before HOT initialises.
	Also ships the site-wide "Manual Save" switch (Excel View Settings)."""
	try:
		bootinfo.excel_formula_configs = _get_active_formula_configs()
	except Exception:
		bootinfo.excel_formula_configs = []
	bootinfo.excel_view_manual_save = _manual_save_enabled()


def _manual_save_enabled() -> int:
	"""Manual Save (Ctrl+S) is OFF unless an administrator switches it on in
	Excel View Settings. Edits are auto-saved, as before, when the setting has never
	been saved (a fresh install or an upgrade)."""
	try:
		stored = frappe.db.sql(
			"select value from tabSingles where doctype=%s and field=%s",
			("Excel View Settings", "manual_save"),
		)
		return int(stored[0][0] or 0) if stored else 0
	except Exception:
		# The settings doctype does not exist yet (before the first migrate).
		return 0


@frappe.whitelist()
def get_formula_configs():
	"""Return active non-system Excel Formula configs as JSON.
	Called client-side after a realtime excel_formula_updated event to
	refresh frappe.boot.excel_formula_configs without a page reload."""
	return _get_active_formula_configs()


def seed_example_formulas():
	"""Insert ready-to-use example formulas with preset filters.
	Run once after migrate: bench --site index.com execute excel_view.api.seed_example_formulas
	"""
	examples = [
		{
			"formula_name": "REVENUE_MTD",
			"label": "Revenue Month-to-Date",
			"category": "Financial",
			"formula_type": "sum",
			"source_doctype": "Sales Invoice",
			"target_fieldname": "grand_total",
			"description": "Total submitted Sales Invoice revenue for the current period.\n=REVENUE_MTD()",
			"is_system": 0,
			"filters": [
				{"filter_key": "docstatus", "filter_value": "1"},
				{"filter_key": "posting_date >=", "filter_value": "PERIOD_START()"},
				{"filter_key": "posting_date <=", "filter_value": "PERIOD_END()"},
			],
		},
		{
			"formula_name": "RECEIVABLES",
			"label": "Total Receivables",
			"category": "Financial",
			"formula_type": "sum",
			"source_doctype": "Sales Invoice",
			"target_fieldname": "outstanding_amount",
			"description": "Sum of all unpaid Sales Invoice outstanding amounts.\n=RECEIVABLES()",
			"is_system": 0,
			"filters": [
				{"filter_key": "docstatus", "filter_value": "1"},
				{"filter_key": "outstanding_amount >", "filter_value": "0"},
			],
		},
		{
			"formula_name": "PAYABLES",
			"label": "Total Payables",
			"category": "Financial",
			"formula_type": "sum",
			"source_doctype": "Purchase Invoice",
			"target_fieldname": "outstanding_amount",
			"description": "Sum of all unpaid Purchase Invoice outstanding amounts.\n=PAYABLES()",
			"is_system": 0,
			"filters": [
				{"filter_key": "docstatus", "filter_value": "1"},
				{"filter_key": "outstanding_amount >", "filter_value": "0"},
			],
		},
		{
			"formula_name": "HEADCOUNT",
			"label": "Active Employee Count",
			"category": "HR",
			"formula_type": "count",
			"source_doctype": "Employee",
			"target_fieldname": "name",
			"description": "Count of currently active employees.\n=HEADCOUNT()",
			"is_system": 0,
			"filters": [{"filter_key": "status", "filter_value": "Active"}],
		},
		{
			"formula_name": "OPEN_ORDERS",
			"label": "Open Sales Orders Value",
			"category": "Sales",
			"formula_type": "sum",
			"source_doctype": "Sales Order",
			"target_fieldname": "grand_total",
			"description": "Total value of open (not fully billed) Sales Orders.\n=OPEN_ORDERS()",
			"is_system": 0,
			"filters": [
				{"filter_key": "docstatus", "filter_value": "1"},
				{"filter_key": "status", "filter_value": "To Bill"},
			],
		},
		{
			"formula_name": "OVERDUE_COUNT",
			"label": "Overdue Invoice Count",
			"category": "Financial",
			"formula_type": "count",
			"source_doctype": "Sales Invoice",
			"target_fieldname": "name",
			"description": "Count of submitted invoices with outstanding amount past due date.\n=OVERDUE_COUNT()",
			"is_system": 0,
			"filters": [
				{"filter_key": "docstatus", "filter_value": "1"},
				{"filter_key": "outstanding_amount >", "filter_value": "0"},
				{"filter_key": "due_date <", "filter_value": "TODAY()"},
			],
		},
		{
			"formula_name": "AVG_INVOICE_VALUE",
			"label": "Average Invoice Value",
			"category": "Sales",
			"formula_type": "avg",
			"source_doctype": "Sales Invoice",
			"target_fieldname": "grand_total",
			"description": "Average Sales Invoice value for the current period.\n=AVG_INVOICE_VALUE()",
			"is_system": 0,
			"filters": [
				{"filter_key": "docstatus", "filter_value": "1"},
				{"filter_key": "posting_date >=", "filter_value": "PERIOD_START()"},
				{"filter_key": "posting_date <=", "filter_value": "PERIOD_END()"},
			],
		},
	]

	inserted = 0
	for ex in examples:
		if frappe.db.exists("Excel Formula", ex["formula_name"]):
			continue
		doc = frappe.new_doc("Excel Formula")
		doc.update({k: v for k, v in ex.items() if k != "filters"})
		doc.is_active = 1
		for f in ex.get("filters", []):
			doc.append("filters", f)
		doc.insert(ignore_permissions=True)
		inserted += 1

	frappe.db.commit()
	print(f"Seeded {inserted} example Excel Formula records.")


# ── V3 IntelliFlow: DuckDB bulk fetch + schema endpoints ──────────────────────


@frappe.whitelist()
def bulk_fetch_for_duckdb(doctype, filters="[]", limit=100000):
	"""
	Bulk-fetch all records of a DocType for client-side DuckDB ingestion.

	Returns Arrow IPC binary (base64) when pyarrow is available — 3-5x smaller
	than JSON and loads directly into DuckDB without CSV re-encoding.
	Falls back to JSON rows for compatibility.

	Filters: JSON-encoded list of [fieldname, operator, value] triples.
	Limit: max rows to fetch (default 100000).
	"""
	frappe.has_permission(doctype, throw=True)
	limit = min(int(limit or 100000), 100000)

	# Parse filters
	parsed_filters = []
	if filters:
		try:
			raw = _fast_parse(filters) if isinstance(filters, str) else (filters or [])
			if isinstance(raw, list):
				parsed_filters = raw
		except Exception:
			pass

	meta = frappe.get_meta(doctype)
	SKIP_TYPES = {
		"Column Break",
		"Section Break",
		"Tab Break",
		"Fold",
		"Heading",
		"HTML",
		"Custom HTML",
		"Table",
		"Table MultiSelect",
		"Password",
	}
	fields = ["name", "creation", "modified", "owner"]
	for df in meta.fields:
		if df.fieldtype in SKIP_TYPES or df.is_virtual:
			continue
		fields.append(df.fieldname)

	# Deduplicate while preserving order
	seen = set()
	unique_fields = []
	for f in fields:
		if f not in seen:
			seen.add(f)
			unique_fields.append(f)

	rows = frappe.get_all(
		doctype,
		fields=unique_fields,
		filters=parsed_filters,
		limit=limit,
		ignore_permissions=False,
	)
	rows = [dict(r) for r in rows]

	# ── Arrow IPC binary transport (3-5x smaller + zero CSV re-encoding) ──
	try:
		import base64

		import pyarrow as pa

		if rows:
			# Build columnar arrays — cast everything to string to avoid type
			# inference issues across Frappe field types (dates, decimals, etc.)
			col_arrays = {}
			for f in unique_fields:
				col_arrays[f] = pa.array(
					[str(r[f]) if r[f] is not None else "" for r in rows], type=pa.string()
				)
			table = pa.table(col_arrays)
		else:
			schema = pa.schema([(f, pa.string()) for f in unique_fields])
			table = pa.table({f: pa.array([], type=pa.string()) for f in unique_fields}, schema=schema)

		sink = pa.BufferOutputStream()
		writer = pa.ipc.new_stream(sink, table.schema)
		writer.write_table(table)
		writer.close()
		arrow_b64 = base64.b64encode(sink.getvalue().to_pybytes()).decode("ascii")

		return {
			"doctype": doctype,
			"fields": unique_fields,
			"count": len(rows),
			"arrow_ipc": arrow_b64,  # base64 Arrow IPC stream
		}

	except Exception:
		# Fallback: plain JSON rows (always works)
		return {
			"doctype": doctype,
			"rows": rows,
			"count": len(rows),
			"fields": unique_fields,
		}


@frappe.whitelist()
def get_doctype_schema(doctypes):
	"""
	Return field metadata for one or more DocTypes — used by QueryFlowPanel
	to render field checklists and EER badges without round-trips.

	Input:  doctypes — JSON list of DocType names, or a single name string.
	Output: {doctype_name: {module, fields: [{fieldname, label, fieldtype, options, reqd}]}}
	"""
	if isinstance(doctypes, str):
		try:
			doctypes = _fast_parse(doctypes) if isinstance(doctypes, str) else (doctypes or [])
		except Exception:
			doctypes = [doctypes]

	if not isinstance(doctypes, list):
		doctypes = [doctypes]

	SKIP_TYPES = {
		"Column Break",
		"Section Break",
		"Tab Break",
		"Fold",
		"Heading",
		"HTML",
		"Custom HTML",
		"Password",
	}

	result = {}
	for doctype in doctypes:
		if not doctype:
			continue
		try:
			frappe.has_permission(doctype, throw=True)
		except frappe.PermissionError:
			continue

		meta = frappe.get_meta(doctype)
		fields = []

		# Always include name first
		fields.append(
			{
				"fieldname": "name",
				"label": "Name (ID)",
				"fieldtype": "Data",
				"options": "",
				"reqd": 1,
			}
		)

		for df in meta.fields:
			if df.fieldtype in SKIP_TYPES:
				continue
			if df.is_virtual:
				continue
			fields.append(
				{
					"fieldname": df.fieldname,
					"label": df.label or df.fieldname,
					"fieldtype": df.fieldtype,
					"options": df.options or "",
					"reqd": int(bool(df.reqd)),
				}
			)

		result[doctype] = {
			"module": meta.module or "",
			"fields": fields,
		}

	return result


# ── Permission Panel API ───────────────────────────────────────────────────────
# All mutations:
#   - @excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
#     replaces the old @frappe.whitelist() + frappe.only_for() + manual checks.
# Read endpoints:
#   - @excel_whitelist(roles=["System Manager"])  (GET is fine for reads)
#
# No direct DB writes — every mutation delegates to Frappe's canonical APIs
# so cache invalidation, validation, and Custom DocPerm copy-on-write all happen
# exactly as they do in Frappe's own Permission Manager page.

_PERM_BLOCKED_DOCTYPES = frozenset(
	{
		"DocType",
		"DocField",
		"DocPerm",
		"Custom DocPerm",
		"Property Setter",
		"Role",
		"User",
		"Session",
		"DefaultValue",
		"Error Log",
		"Access Log",
	}
)


def _perm_guard(doctype):
	"""Block meta-system doctypes. Role check is handled by @excel_whitelist."""
	if doctype in _PERM_BLOCKED_DOCTYPES:
		frappe.throw(
			f"Permissions for system DocType '{frappe.bold(doctype)}' cannot be " "managed from Excel View.",
			frappe.PermissionError,
		)


@excel_whitelist(roles=["System Manager"])
def get_doctype_permissions(doctype):
	"""
	Return active permission rows for the DocType.
	Delegates to permission_manager.get_permissions() which automatically uses
	Custom DocPerm when present, Standard DocPerm otherwise.
	"""
	_perm_guard(doctype)
	from frappe.core.page.permission_manager import permission_manager as pm

	return pm.get_permissions(doctype=doctype)


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def update_role_permission(doctype, role, permlevel, ptype, value=None, if_owner=0):
	"""
	Toggle a single permission bit.
	Delegates to permission_manager.update() which calls update_permission_property(),
	validates, and schedules cache clear after DB commit.
	Returns "refresh" if Custom DocPerm was newly initialised, None otherwise.
	"""
	_perm_guard(doctype)
	from frappe.core.page.permission_manager import permission_manager as pm

	return pm.update(
		doctype=doctype,
		role=role,
		permlevel=frappe.utils.cint(permlevel),
		ptype=ptype,
		value=value,
		if_owner=if_owner,
	)


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def add_role_permission(doctype, role, permlevel=0):
	"""
	Add a new (role, permlevel) row via frappe.permissions.add_permission.
	That function calls setup_custom_perms() then ORM-saves the new row,
	so on_update() fires and clears doctype cache automatically.
	"""
	_perm_guard(doctype)
	from frappe.permissions import add_permission

	add_permission(doctype, role, frappe.utils.cint(permlevel))


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def remove_role_permission(doctype, role, permlevel, if_owner=0):
	"""
	Remove a (role, permlevel) row.
	Delegates to permission_manager.remove() which validates that at least
	one permission row remains and calls validate_permissions_for_doctype.
	"""
	_perm_guard(doctype)
	from frappe.core.page.permission_manager import permission_manager as pm

	pm.remove(
		doctype=doctype,
		role=role,
		permlevel=frappe.utils.cint(permlevel),
		if_owner=if_owner,
	)


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def reset_doctype_permissions(doctype):
	"""
	Delete all Custom DocPerm and revert to the app's Standard DocPerm.
	Delegates to permission_manager.reset() which calls reset_perms()
	then clear_permissions_cache().
	"""
	_perm_guard(doctype)
	from frappe.core.page.permission_manager import permission_manager as pm

	pm.reset(doctype=doctype)


@excel_whitelist(roles=["System Manager"])
def get_field_permlevels(doctype):
	"""
	Return every data field with its current permlevel.
	frappe.get_meta() already merges Property Setter overrides, so df.permlevel
	is always the live value — no separate Property Setter query needed.
	"""
	_perm_guard(doctype)
	SKIP_TYPES = frozenset(
		{
			"Section Break",
			"Column Break",
			"Tab Break",
			"Fold",
			"Heading",
			"HTML",
			"Custom HTML",
			"Password",
			"Table",
			"Table MultiSelect",
		}
	)
	meta = frappe.get_meta(doctype)
	return [
		{
			"fieldname": df.fieldname,
			"label": df.label or df.fieldname,
			"fieldtype": df.fieldtype,
			"permlevel": frappe.utils.cint(df.permlevel),
		}
		for df in meta.fields
		if df.fieldtype not in SKIP_TYPES and df.fieldname and not df.is_virtual
	]


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def set_field_permlevel(doctype, fieldname, permlevel):
	"""
	Change a field's permlevel via Property Setter (upgrade-safe).
	frappe.make_property_setter creates/updates a Property Setter record and
	calls frappe.clear_cache(doctype) inside its validate() hook automatically.
	"""
	_perm_guard(doctype)
	permlevel = frappe.utils.cint(permlevel)
	if not 0 <= permlevel <= 9:
		frappe.throw("permlevel must be between 0 and 9.")

	# Validate field exists in this doctype
	meta = frappe.get_meta(doctype)
	if not meta.get_field(fieldname):
		frappe.throw(f"Field '{frappe.bold(fieldname)}' not found in {frappe.bold(doctype)}.")

	frappe.make_property_setter(
		{
			"doctype": doctype,
			"doctype_or_field": "DocField",
			"fieldname": fieldname,
			"property": "permlevel",
			"value": permlevel,
			"property_type": "Int",
		},
		is_system_generated=False,
	)

	# Belt-and-suspenders: schedule one more clear after commit
	frappe.db.after_commit.add(lambda: frappe.clear_cache(doctype=doctype))


@excel_whitelist(roles=["System Manager"])
def get_role_user_counts():
	"""
	Return {role_name: user_count} for every role that has at least one user.
	Single aggregating query — O(1) round-trips regardless of role count.
	"""
	rows = frappe.db.sql(
		"SELECT role, COUNT(*) AS cnt" " FROM `tabHas Role`" " WHERE parenttype = 'User'" " GROUP BY role",
		as_dict=True,
	)
	return {r.role: r.cnt for r in rows}


@excel_whitelist(roles=["System Manager"])
def get_all_roles():
	"""
	Return sorted list of non-disabled, non-automatic role names.
	Mirrors permission_manager.get_roles_and_doctypes() role list.
	"""
	try:
		from frappe.permissions import AUTOMATIC_ROLES

		restricted = set(AUTOMATIC_ROLES)
	except ImportError:
		restricted = {"All", "Guest"}

	if frappe.session.user != "Administrator":
		custom_user_type_roles = frappe.get_all(
			"User Type", filters={"is_standard": 0}, fields=["role"], pluck="role"
		)
		restricted.update(custom_user_type_roles)

	restricted.add("Administrator")

	return frappe.get_all(
		"Role",
		filters={"disabled": 0, "name": ["not in", list(restricted)]},
		pluck="name",
		order_by="name asc",
	)


# ── V3.4.4 — Access Profiles (Role Profiles + Module Profiles) ────────────────


@excel_whitelist(roles=["System Manager"])
def get_access_profiles():
	"""
	Single round-trip for the Access Profiles tab.
	4 lightweight queries merged in Python — O(profiles + assignments + modules).
	Returns all role profiles, module profiles, and installed module names.
	"""
	# ── Role Profiles ──────────────────────────────────────────────────────
	rp_rows = frappe.get_all("Role Profile", fields=["name"], order_by="name asc")

	if rp_rows:
		rp_counts = {
			r.role_profile_name: r.cnt
			for r in frappe.db.sql(
				"SELECT role_profile_name, COUNT(*) AS cnt"
				" FROM `tabUser`"
				" WHERE role_profile_name IS NOT NULL AND role_profile_name != ''"
				"   AND enabled = 1"
				" GROUP BY role_profile_name",
				as_dict=True,
			)
		}
		rp_role_map = {}
		for r in frappe.get_all(
			"Has Role",
			filters={"parenttype": "Role Profile"},
			fields=["parent", "role"],
		):
			rp_role_map.setdefault(r.parent, []).append(r.role)
	else:
		rp_counts, rp_role_map = {}, {}

	role_profiles = [
		{
			"name": rp.name,
			"user_count": rp_counts.get(rp.name, 0),
			"roles": rp_role_map.get(rp.name, []),
		}
		for rp in rp_rows
	]

	# ── Module Profiles ────────────────────────────────────────────────────
	mp_rows = frappe.get_all("Module Profile", fields=["name"], order_by="name asc")

	if mp_rows:
		mp_counts = {
			r.module_profile: r.cnt
			for r in frappe.db.sql(
				"SELECT module_profile, COUNT(*) AS cnt"
				" FROM `tabUser`"
				" WHERE module_profile IS NOT NULL AND module_profile != ''"
				"   AND enabled = 1"
				" GROUP BY module_profile",
				as_dict=True,
			)
		}
		mp_mod_map = {}
		for r in frappe.get_all(
			"Block Module",
			filters={"parenttype": "Module Profile"},
			fields=["parent", "module"],
		):
			mp_mod_map.setdefault(r.parent, []).append(r.module)
	else:
		mp_counts, mp_mod_map = {}, {}

	module_profiles = [
		{
			"name": mp.name,
			"user_count": mp_counts.get(mp.name, 0),
			"blocked_modules": mp_mod_map.get(mp.name, []),
		}
		for mp in mp_rows
	]

	# ── All installed modules ──────────────────────────────────────────────
	from frappe.config import get_modules_from_all_apps

	all_modules = sorted(m["module_name"] for m in get_modules_from_all_apps() if m.get("module_name"))

	return {
		"role_profiles": role_profiles,
		"module_profiles": module_profiles,
		"all_modules": all_modules,
	}


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def save_role_profile(profile_name, roles):
	"""
	Replace the full role list for a Role Profile in one ORM save.
	Frappe's on_update() queues user-sync automatically after commit.
	roles: JSON-encoded list of role name strings.
	"""
	import json as _json

	role_list = _fast_parse(roles) if isinstance(roles, str) else list(roles)

	doc = frappe.get_doc("Role Profile", profile_name)
	doc.roles = []
	for role in role_list:
		doc.append("roles", {"role": role})
	doc.save(ignore_permissions=True)


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def save_module_profile(profile_name, blocked_modules):
	"""
	Replace the full blocked_modules list for a Module Profile in one ORM save.
	Frappe's on_update() queues user-sync automatically after commit.
	blocked_modules: JSON-encoded list of module name strings.
	"""
	import json as _json

	mod_list = _fast_parse(blocked_modules) if isinstance(blocked_modules, str) else list(blocked_modules)

	doc = frappe.get_doc("Module Profile", profile_name)
	doc.block_modules = []
	for module in mod_list:
		doc.append("block_modules", {"module": module})
	doc.save(ignore_permissions=True)


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def create_access_profile(profile_type, name):
	"""
	Create a new Role Profile or Module Profile.
	profile_type: "role" | "module"
	"""
	name = (name or "").strip()
	if not name:
		frappe.throw(_("Profile name cannot be empty."))

	if profile_type == "role":
		if frappe.db.exists("Role Profile", name):
			frappe.throw(_("Role Profile {0} already exists.").format(frappe.bold(name)))
		doc = frappe.new_doc("Role Profile")
		doc.role_profile = name
		doc.insert(ignore_permissions=True)
	elif profile_type == "module":
		if frappe.db.exists("Module Profile", name):
			frappe.throw(_("Module Profile {0} already exists.").format(frappe.bold(name)))
		doc = frappe.new_doc("Module Profile")
		doc.module_profile_name = name
		doc.insert(ignore_permissions=True)
	else:
		frappe.throw(_("Invalid profile_type."))

	return {"name": doc.name}


@excel_whitelist(roles=["System Manager"], methods=["POST"], audit=True)
def delete_access_profile(profile_type, name):
	"""Delete a Role Profile or Module Profile."""
	if profile_type == "role":
		frappe.delete_doc("Role Profile", name, ignore_permissions=True)
	elif profile_type == "module":
		frappe.delete_doc("Module Profile", name, ignore_permissions=True)
	else:
		frappe.throw(_("Invalid profile_type."))
